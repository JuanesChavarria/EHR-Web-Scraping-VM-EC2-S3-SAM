from fastapi import FastAPI
app = FastAPI()

@app.get("/GetPatientUpserts")
async def get_patient():
    return {
        "message": "Success",
        "loader": {
            "name": "Appointment and Patients Navigation Site",
            "version": "1.0.0",
            "status": "completed",
            "files": [
                {
                    "name": "PatientInserts",
                    "url": "s3url/patients/patient_inserts.csv",
                },
                {
                    "name": "PatientUpdates",
                    "url": "s3url/patients/patients_updates.csv",
                }
            ]

        }
    }

@app.get("/GetServiceAppointment")
async def get_service_appointment():
    return {
        "message": "TEST 1",
    }
