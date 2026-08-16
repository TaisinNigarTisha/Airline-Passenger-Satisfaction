import React, { useState } from 'react'

const API_URL = 'http://127.0.0.1:8000/predict'

const RATING_FIELDS = [
  { key: 'Departure and Arrival Time Convenience', label: 'Departure & Arrival Time Convenience' },
  { key: 'Ease of Online Booking', label: 'Ease of Online Booking' },
  { key: 'Check-in Service', label: 'Check-in Service' },
  { key: 'Online Boarding', label: 'Online Boarding' },
  { key: 'Gate Location', label: 'Gate Location' },
  { key: 'On-board Service', label: 'On-board Service' },
  { key: 'Seat Comfort', label: 'Seat Comfort' },
  { key: 'Leg Room Service', label: 'Leg Room Service' },
  { key: 'Cleanliness', label: 'Cleanliness' },
  { key: 'Food and Drink', label: 'Food and Drink' },
  { key: 'In-flight Service', label: 'In-flight Service' },
  { key: 'In-flight Wifi Service', label: 'In-flight Wifi Service' },
  { key: 'In-flight Entertainment', label: 'In-flight Entertainment' },
  { key: 'Baggage Handling', label: 'Baggage Handling' },
]

const DEFAULT_FORM = {
  Gender: 'Male',
  Age: 35,
  'Customer Type': 'Returning',
  'Type of Travel': 'Business',
  Class: 'Business',
  'Flight Distance': 1200,
  'Departure Delay': 0,
  'Arrival Delay': 0,
  ...Object.fromEntries(RATING_FIELDS.map((f) => [f.key, 3])),
}

function RatingSlider({ label, value, onChange }) {
  return (
    <div className="rating-row">
      <label>
        {label}
        <span className="rating-value">{value === 0 ? 'N/A' : value}</span>
      </label>
      <input
        type="range"
        min="0"
        max="5"
        step="1"
        value={value}
        onChange={(e) => onChange(Number(e.target.value))}
      />
      <div className="rating-ticks">
        <span>N/A</span>
        <span>1</span>
        <span>2</span>
        <span>3</span>
        <span>4</span>
        <span>5</span>
      </div>
    </div>
  )
}

export default function App() {
  const [form, setForm] = useState(DEFAULT_FORM)
  const [result, setResult] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  const updateField = (key, value) => {
    setForm((prev) => ({ ...prev, [key]: value }))
  }

  const validate = () => {
    if (form.Age < 0 || form.Age > 120) return 'Age must be between 0 and 120.'
    if (form['Flight Distance'] < 0) return 'Flight distance cannot be negative.'
    if (form['Departure Delay'] < 0 || form['Arrival Delay'] < 0)
      return 'Delay values cannot be negative.'
    return null
  }

  const handleSubmit = async (e) => {
    e.preventDefault()
    setError(null)
    setResult(null)

    const validationError = validate()
    if (validationError) {
      setError(validationError)
      return
    }

    setLoading(true)
    try {
      const response = await fetch(API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      })

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}))
        throw new Error(errBody.detail || `Request failed with status ${response.status}`)
      }

      const data = await response.json()
      setResult(data)
    } catch (err) {
      setError(
        err.message === 'Failed to fetch'
          ? 'Could not reach the prediction server. Make sure the FastAPI backend is running on http://127.0.0.1:8000.'
          : err.message
      )
    } finally {
      setLoading(false)
    }
  }

  const handleReset = () => {
    setForm(DEFAULT_FORM)
    setResult(null)
    setError(null)
  }

  return (
    <div className="page">
      <header className="header">
        <h1>Service Optimization</h1>
        <p>Airline Passenger Satisfaction Predictor</p>
      </header>

      <main className="main">
        <form className="card form" onSubmit={handleSubmit}>
          <h2>Passenger &amp; Trip Details</h2>
          <div className="grid-2">
            <div className="field">
              <label>Gender</label>
              <select value={form.Gender} onChange={(e) => updateField('Gender', e.target.value)}>
                <option value="Male">Male</option>
                <option value="Female">Female</option>
              </select>
            </div>

            <div className="field">
              <label>Age</label>
              <input
                type="number"
                min="0"
                max="120"
                value={form.Age}
                onChange={(e) => updateField('Age', Number(e.target.value))}
              />
            </div>

            <div className="field">
              <label>Customer Type</label>
              <select
                value={form['Customer Type']}
                onChange={(e) => updateField('Customer Type', e.target.value)}
              >
                <option value="Returning">Returning</option>
                <option value="First-time">First-time</option>
              </select>
            </div>

            <div className="field">
              <label>Type of Travel</label>
              <select
                value={form['Type of Travel']}
                onChange={(e) => updateField('Type of Travel', e.target.value)}
              >
                <option value="Business">Business</option>
                <option value="Personal">Personal</option>
              </select>
            </div>

            <div className="field">
              <label>Class</label>
              <select value={form.Class} onChange={(e) => updateField('Class', e.target.value)}>
                <option value="Business">Business</option>
                <option value="Economy">Economy</option>
                <option value="Economy Plus">Economy Plus</option>
              </select>
            </div>

            <div className="field">
              <label>Flight Distance (miles)</label>
              <input
                type="number"
                min="0"
                value={form['Flight Distance']}
                onChange={(e) => updateField('Flight Distance', Number(e.target.value))}
              />
            </div>

            <div className="field">
              <label>Departure Delay (minutes)</label>
              <input
                type="number"
                min="0"
                value={form['Departure Delay']}
                onChange={(e) => updateField('Departure Delay', Number(e.target.value))}
              />
            </div>

            <div className="field">
              <label>Arrival Delay (minutes)</label>
              <input
                type="number"
                min="0"
                value={form['Arrival Delay']}
                onChange={(e) => updateField('Arrival Delay', Number(e.target.value))}
              />
            </div>
          </div>

          <h2>Service Ratings</h2>
          <p className="hint">Slide from N/A (not applicable) up to 5 (excellent).</p>
          <div className="ratings-grid">
            {RATING_FIELDS.map((f) => (
              <RatingSlider
                key={f.key}
                label={f.label}
                value={form[f.key]}
                onChange={(v) => updateField(f.key, v)}
              />
            ))}
          </div>

          {error && <div className="alert error">{error}</div>}

          <div className="actions">
            <button type="submit" className="btn primary" disabled={loading}>
              {loading ? 'Predicting…' : 'Predict Satisfaction'}
            </button>
            <button type="button" className="btn secondary" onClick={handleReset} disabled={loading}>
              Reset
            </button>
          </div>
        </form>

        <div className="card result-card">
          <h2>Prediction Result</h2>
          {!result && !loading && (
            <p className="placeholder">Fill in the form and click "Predict Satisfaction" to see a result.</p>
          )}
          {loading && <p className="placeholder">Contacting the prediction model…</p>}
          {result && (
            <div className={`result ${result.prediction === 'Satisfied' ? 'positive' : 'negative'}`}>
              <p className="result-label">Passenger is likely to be:</p>
              <p className="result-value">{result.prediction.toUpperCase()}</p>
              <p className="result-confidence">
                Confidence: {(result.confidence * 100).toFixed(1)}%
              </p>
              <div className="prob-bars">
                {Object.entries(result.probabilities).map(([label, prob]) => (
                  <div key={label} className="prob-row">
                    <span>{label}</span>
                    <div className="prob-bar-bg">
                      <div className="prob-bar-fill" style={{ width: `${prob * 100}%` }} />
                    </div>
                    <span>{(prob * 100).toFixed(1)}%</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </main>

      <footer className="footer">
        <p>Data Mining Lab Project — Service Optimization Using Airline Passenger Satisfaction Data</p>
      </footer>
    </div>
  )
}
