from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware  # type: ignore
import pandas as pd
import io
import numpy as np

import math

app = FastAPI()


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def calculate_percentage_change(df, column_name):
    df['percentage_change'] = df[column_name].pct_change() * 100
    df['percentage_change'] = df['percentage_change'].where(pd.notna(df['percentage_change']), 0)
    return df[['Month', 'percentage_change']].to_dict(orient='records')


@app.post("/uploadfile/")
async def create_upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        return JSONResponse(status_code=400, content={"message": "Invalid file type. Please upload an Excel file."})

    try:
        contents = await file.read()
        xls = pd.ExcelFile(io.BytesIO(contents))

        results = {}

        # Carbon Emissions
        if 'Carbon Emissions' in xls.sheet_names:
            df_carbon = pd.read_excel(xls, 'Carbon Emissions')
            results['carbon_emissions_change'] = calculate_percentage_change(df_carbon, 'Carbon Emissions (tons)')

        # Water Usage
        if 'Water Usage' in xls.sheet_names:
            df_water = pd.read_excel(xls, 'Water Usage')
            results['water_usage_change'] = calculate_percentage_change(df_water, 'Water Usage (cubic meters)')

        # Employee Safety

        # Energy Sources
        if 'Energy Sources' in xls.sheet_names:
            df_energy = pd.read_excel(xls, 'Energy Sources')
            combined_energy = {
                'renewable': df_energy['Renewable (%)'].mean().item(),
                'non_renewable': df_energy['Non-Renewable (%)'].mean().item()
            }
            last_month_energy = {
                'renewable': df_energy.iloc[-1]['Renewable (%)'].item(),
                'non_renewable': df_energy.iloc[-1]['Non-Renewable (%)'].item()
            }
            results['energy_sources'] = {
                'combined': combined_energy,
                'last_month': last_month_energy
            }

        # Waste Management
        if 'Waste Management' in xls.sheet_names:
            df_waste = pd.read_excel(xls, 'Waste Management')
            combined_waste = {
                'recycled': df_waste['Recycled (%)'].mean().item(),
                'unrecycled': df_waste['Unrecycled (%)'].mean().item()
            }
            last_month_waste = {
                'recycled': df_waste.iloc[-1]['Recycled (%)'].item(),
                'unrecycled': df_waste.iloc[-1]['Unrecycled (%)'].item()
            }
            results['waste_management'] = {
                'combined': combined_waste,
                'last_month': last_month_waste
            }

        for key, val in results.items():
            for item in results[key]:
                if isinstance(item, dict):
                    if item["percentage_change"] == (float('inf') or float('-inf')):
                        item["percentage_change"] = None

        if 'Employee Safety' in xls.sheet_names:
            df_safety = pd.read_excel(xls, 'Employee Safety')
            results['employee_safety_change'] = df_safety.to_dict(orient='records')


        return results

    except Exception as e:
        return JSONResponse(status_code=500, content={"message": f"There was an error processing the file: {e}"})