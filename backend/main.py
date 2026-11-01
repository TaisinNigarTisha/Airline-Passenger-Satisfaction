"""
Service Optimization — Airline Passenger Satisfaction Predictor
FastAPI backend

Loads the trained scikit-learn Pipeline (preprocessing + Random Forest classifier)
saved by the notebook (Section 17) and exposes a single POST /predict endpoint.

Run with:
    uvicorn main:app --reload
"""

import os

# Load environment variables from .env file if available
if os.path.exists(".env"):
    with open(".env") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#"):
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip())

# Read allowed CORS origins from environment (comma-separated)
ALLOWED_ORIGINS = os.getenv(
    "ALLOWED_ORIGINS",
    "http://localhost:5173,http://127.0.0.1:5173,http://localhost:3000,http://127.0.0.1:3000",
).split(",")

from huggingface_hub import hf_hub_download
import joblib
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from typing import Literal

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = FastAPI(
    title="Service Optimization Prediction API",
    description="Predicts airline passenger satisfaction from passenger and service data.",
    version="1.0.0",
)

# Allow the React dev server (Vite default port 5173, CRA default 3000) to call this API.
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------------------------
# Load the trained model pipeline at startup from Hugging Face Hub
# ---------------------------------------------------------------------------
REPO_ID = os.getenv("REPO_ID", "taisinnigartisha/Airline_Passenger_Satisfaction_Predictor")
FILENAME = os.getenv("FILENAME", "sklearn_model.joblib")

model = None
model_load_error = None
try:
    model_path = hf_hub_download(repo_id=REPO_ID, filename=FILENAME)
    model = joblib.load(model_path)
except Exception as exc:  # noqa: BLE001
    model_load_error = str(exc)

# The label encoder classes are alphabetical: ['Neutral or Dissatisfied', 'Satisfied']
# (matches sklearn.preprocessing.LabelEncoder behaviour used during training)
CLASS_LABELS = ["Neutral or Dissatisfied", "Satisfied"]


# ---------------------------------------------------------------------------
# Request schema — mirrors the exact feature names/types used to train the model
# ---------------------------------------------------------------------------
class PassengerInput(BaseModel):
    Gender: Literal["Male", "Female"]
    Age: int = Field(..., ge=0, le=120)
    Customer_Type: Literal["First-time", "Returning"] = Field(..., alias="Customer Type")
    Type_of_Travel: Literal["Business", "Personal"] = Field(..., alias="Type of Travel")
    Class: Literal["Business", "Economy", "Economy Plus"]
    Flight_Distance: int = Field(..., ge=0, alias="Flight Distance")
    Departure_Delay: int = Field(..., ge=0, alias="Departure Delay")
    Arrival_Delay: float = Field(..., ge=0, alias="Arrival Delay")

    Departure_and_Arrival_Time_Convenience: int = Field(
        ..., ge=0, le=5, alias="Departure and Arrival Time Convenience"
    )
    Ease_of_Online_Booking: int = Field(..., ge=0, le=5, alias="Ease of Online Booking")
    Check_in_Service: int = Field(..., ge=0, le=5, alias="Check-in Service")
    Online_Boarding: int = Field(..., ge=0, le=5, alias="Online Boarding")
    Gate_Location: int = Field(..., ge=0, le=5, alias="Gate Location")
    On_board_Service: int = Field(..., ge=0, le=5, alias="On-board Service")
    Seat_Comfort: int = Field(..., ge=0, le=5, alias="Seat Comfort")
    Leg_Room_Service: int = Field(..., ge=0, le=5, alias="Leg Room Service")
    Cleanliness: int = Field(..., ge=0, le=5)
    Food_and_Drink: int = Field(..., ge=0, le=5, alias="Food and Drink")
    In_flight_Service: int = Field(..., ge=0, le=5, alias="In-flight Service")
    In_flight_Wifi_Service: int = Field(..., ge=0, le=5, alias="In-flight Wifi Service")
    In_flight_Entertainment: int = Field(..., ge=0, le=5, alias="In-flight Entertainment")
    Baggage_Handling: int = Field(..., ge=0, le=5, alias="Baggage Handling")

    model_config = {
        "populate_by_name": True,
        "json_schema_extra": {
            "example": {
                "Gender": "Male",
                "Age": 35,
                "Customer Type": "Returning",
                "Type of Travel": "Business",
                "Class": "Business",
                "Flight Distance": 1200,
                "Departure Delay": 10,
                "Arrival Delay": 5,
                "Departure and Arrival Time Convenience": 4,
                "Ease of Online Booking": 4,
                "Check-in Service": 5,
                "Online Boarding": 4,
                "Gate Location": 3,
                "On-board Service": 5,
                "Seat Comfort": 5,
                "Leg Room Service": 4,
                "Cleanliness": 5,
                "Food and Drink": 4,
                "In-flight Service": 5,
                "In-flight Wifi Service": 4,
                "In-flight Entertainment": 5,
                "Baggage Handling": 5,
            }
        },
    }


class PredictionResponse(BaseModel):
    prediction: str
    confidence: float
    probabilities: dict


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.get("/")
def root():
    return {
        "message": "Service Optimization Prediction API is running.",
        "model_loaded": model is not None,
        "docs": "/docs",
    }


@app.get("/health")
def health():
    if model is None:
        raise HTTPException(status_code=503, detail=f"Model not loaded: {model_load_error}")
    return {"status": "ok"}


@app.post("/predict", response_model=PredictionResponse)
def predict(passenger: PassengerInput):
    if model is None:
        raise HTTPException(
            status_code=503,
            detail=f"Model is not loaded on the server: {model_load_error}",
        )

    # Build a single-row DataFrame with the EXACT column names the pipeline was trained on.
    row = {
        "Gender": passenger.Gender,
        "Age": passenger.Age,
        "Customer Type": passenger.Customer_Type,
        "Type of Travel": passenger.Type_of_Travel,
        "Class": passenger.Class,
        "Flight Distance": passenger.Flight_Distance,
        "Departure Delay": passenger.Departure_Delay,
        "Arrival Delay": passenger.Arrival_Delay,
        "Departure and Arrival Time Convenience": passenger.Departure_and_Arrival_Time_Convenience,
        "Ease of Online Booking": passenger.Ease_of_Online_Booking,
        "Check-in Service": passenger.Check_in_Service,
        "Online Boarding": passenger.Online_Boarding,
        "Gate Location": passenger.Gate_Location,
        "On-board Service": passenger.On_board_Service,
        "Seat Comfort": passenger.Seat_Comfort,
        "Leg Room Service": passenger.Leg_Room_Service,
        "Cleanliness": passenger.Cleanliness,
        "Food and Drink": passenger.Food_and_Drink,
        "In-flight Service": passenger.In_flight_Service,
        "In-flight Wifi Service": passenger.In_flight_Wifi_Service,
        "In-flight Entertainment": passenger.In_flight_Entertainment,
        "Baggage Handling": passenger.Baggage_Handling,
    }

    try:
        X = pd.DataFrame([row])
        pred_encoded = model.predict(X)[0]
        proba = model.predict_proba(X)[0]
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"Prediction failed: {exc}") from exc

    predicted_label = CLASS_LABELS[int(pred_encoded)]
    probabilities = {CLASS_LABELS[i]: round(float(p), 4) for i, p in enumerate(proba)}
    confidence = round(float(max(proba)), 4)

    return PredictionResponse(
        prediction=predicted_label,
        confidence=confidence,
        probabilities=probabilities,
    )
