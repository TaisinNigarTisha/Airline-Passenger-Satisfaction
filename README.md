# Service Optimization Using Airline Passenger Satisfaction Data

A complete Data Mining Lab project: dataset → Jupyter Notebook (preprocessing, EDA, feature
engineering, classification, clustering, service-optimization analysis) → saved model →
FastAPI backend → ReactJS prediction UI.

## Project structure

```
service-optimization/
├── dataset/
│   ├── airline_passenger_satisfaction.csv
│   └── data_dictionary.csv
├── notebook/
│   └── service_optimization.ipynb      # full data mining pipeline, Sections 1–18
├── model/
│   └── service_optimization_model.pkl  # trained pipeline (preprocessing + tuned Random Forest)
├── backend/
│   ├── main.py                         # FastAPI app with POST /predict
│   ├── requirements.txt
│   └── model/
│       └── service_optimization_model.pkl
└── frontend/                           # ReactJS (Vite) prediction UI
    ├── index.html
    ├── package.json
    ├── vite.config.js
    └── src/
        ├── main.jsx
        ├── App.jsx
        └── index.css
```

## Results summary (from the real dataset — 129,880 rows, no fabricated numbers)

| Model | Accuracy | Precision | Recall | F1 | ROC-AUC |
|---|---|---|---|---|---|
| Decision Tree | 0.9501 | 0.9581 | 0.9257 | 0.9416 | 0.9829 |
| Naive Bayes | 0.8626 | 0.8592 | 0.8178 | 0.8380 | 0.9228 |
| Random Forest | 0.9594 | 0.9647 | 0.9410 | 0.9527 | 0.9931 |
| SVM (8,000-row subsample) | 0.9379 | 0.9365 | 0.9193 | 0.9278 | 0.9811 |
| **Random Forest (tuned, final model)** | **0.9632** | **0.9731** | **0.9414** | **0.9570** | **0.9941** |

K-Means clustering (k=2, chosen by silhouette score) found a clean split: a
"positive service experience" cluster (64.2% satisfied) and a "poor service experience"
cluster (82.5% dissatisfied) — full details and service-optimization recommendations are in
Section 16 of the notebook.

## How to run everything end-to-end

### 1. Notebook (optional — the model is already trained and saved)

```bash
cd notebook
jupyter notebook service_optimization.ipynb
```

Run all cells top to bottom. It will overwrite `model/service_optimization_model.pkl` with a
freshly trained model (should reproduce the same results above, since a fixed `random_state=42`
is used throughout).

### 2. Backend (FastAPI)

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload
```

- API runs at `http://127.0.0.1:8000`
- Interactive docs (Swagger UI): `http://127.0.0.1:8000/docs`
- Test the `/predict` endpoint there with the example request body shown in the docs.

### 3. Frontend (ReactJS)

In a new terminal:

```bash
cd frontend
npm install
npm run dev
```

- Open `http://localhost:5173` in your browser.
- Fill in the passenger and service-rating form, click **Predict Satisfaction**.
- The React app calls `POST http://127.0.0.1:8000/predict` on the running FastAPI backend and
  displays the predicted label (`Satisfied` / `Neutral or Dissatisfied`) with confidence.

### Testing checklist

1. Notebook runs top-to-bottom without errors → produces `model/service_optimization_model.pkl`.
2. `uvicorn main:app --reload` starts without errors; `http://127.0.0.1:8000/health` returns `{"status":"ok"}`.
3. `http://127.0.0.1:8000/docs` → try `/predict` with the example payload → returns a `prediction` field.
4. `npm run dev` starts the React app on port 5173.
5. Fill the form, click Predict → result card shows a prediction and probability bars.
6. Try Reset → form returns to defaults, result clears.
7. Stop the backend and click Predict again → the UI shows a clear "could not reach the prediction server" error instead of crashing.
