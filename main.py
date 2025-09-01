from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware  # type: ignore
import pandas as pd
import io
import numpy as np

import math

from google import genai

import os

from dotenv import load_dotenv


load_dotenv()



app = FastAPI()

client = genai.Client(api_key=os.getenv("GENAI_API_KEY"))



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
    return df[['Year', 'percentage_change']].to_dict(orient='records')


def labor_rights_compliance_score(df):
    safety_weighted = (
        df["Accident Fatal"].sum() * 10
        + df["Accident Serious"].sum() * 3
        + df["Accident Minor"].sum() * 1
    )

    # scale by total employees (approx using last row of bargaining employees as proxy)
    total_employees = df["Employees covered by collective bargaining Persons annual"].iloc[-1]
    safety_score = max(0, 100 - (safety_weighted / total_employees) * 1000)  # scaling factor 1000 chosen

    # ---- Diversity Score ----
    board_women = 20        # from your data
    management_women = 61.6 # from your data
    board_score = 100 - abs(50 - board_women)
    management_score = 100 - abs(50 - management_women)
    diversity_score = (board_score + management_score) / 2

    # ---- Wellness Score ----
    covered = df["Employees covered by collective bargaining Persons annual"].iloc[-1]
    wellness_score = (covered / total_employees) * 100

    # ---- Turnover Score ----
    turnover = df["Voluntary Employee Turnover Rate  % annual"].iloc[-1]
    turnover_score = 100 - abs(turnover - 5)

    # ---- Final Compliance Score ----
    compliance_score = (
        0.4 * safety_score +
        0.2 * diversity_score +
        0.2 * wellness_score +
        0.2 * turnover_score
    )

    return int(compliance_score)



@app.post("/uploadfile/")
async def create_upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        return JSONResponse(status_code=400, content={"message": "Invalid file type. Please upload an Excel file."})

    try:
        contents = await file.read()
        xls = pd.ExcelFile(io.BytesIO(contents))

        results = {}
        df = pd.read_excel(xls, 'ESG Metrics')
        results['water_usage_change'] = calculate_percentage_change(df, 'Water Usage (m3)')
        results['water_usage'] = df[["Year", "Water Usage (m3)"]].to_dict(orient='records')

        results['employee_safety_change'] = calculate_percentage_change(df, 'Employee Safety (accidents)')

        results['waste_recycled_change'] = calculate_percentage_change(df, 'Waste Recycled (tons)')
        results['waste_unrecycled_change'] = calculate_percentage_change(df, 'Waste Unrecycled (tons)')
        waste_data = df[["Year", "Waste Recycled (tons)", "Waste Unrecycled (tons)"]].to_dict(orient='records')
        results['waste_recycled'] = waste_data
        latest_waste = waste_data[-1]  # assumes chronological order
        results['waste_recycled_latest'] = [
            {"name": "Recycled", "value": latest_waste["Waste Recycled (tons)"], "color": "#10B981"},
            {"name": "Unrecycled", "value": latest_waste["Waste Unrecycled (tons)"], "color": "#F59E0B"}
        ]

        results['carbon_emissions_change'] = calculate_percentage_change(df, 'Carbon Emissions (tons CO2e)')
        results['carbon_emissions_renewable_change'] = calculate_percentage_change(df, 'Carbon Emissions Renewable (%)')
        results['carbon_emissions_nonrenewable_change'] = calculate_percentage_change(df, 'Carbon Emissions Non-Renewable (%)')
        results['carbon_emissions'] = df[["Year", "Carbon Emissions Renewable (%)", "Carbon Emissions Non-Renewable (%)"]].to_dict(orient='records')

        energy_data = df[["Year", "Energy Renewable (%)", "Energy Non-Renewable (%)"]].to_dict(orient='records')
        results['energy_usage'] = energy_data

        latest_energy = energy_data[-1]  # assumes chronological order
        results['energy_usage_latest'] = [
            {"name": "Renewable", "value": latest_energy["Energy Renewable (%)"], "color": "#10B981"},
            {"name": "Non-Renewable", "value": latest_energy["Energy Non-Renewable (%)"], "color": "#EF4444"}
        ]

        results['energy_renewable_change'] = calculate_percentage_change(df, 'Energy Renewable (%)')
        results['energy_nonrenewable_change'] = calculate_percentage_change(df, 'Energy Non-Renewable (%)')

        results['safety_data'] = [
            {"type" : "Fatal", "count": int(df["Accident Fatal"].sum()), "color": "#EF4444"},
            {"type" : "Serious", "count": int(df[["Accident Serious"]].sum()), "color": "#F59E0B"},
            {"type" : "Minor", "count": int(df[["Accident Minor"]].sum()), "color": "#10B981"},
        ]

        results['diversity_data'] = [
            {
                "category": "Board",
                "men": df['Board members (male) as % of total'].tail(1).values[0],
                "women": df['Board members (female) as % of total'].tail(1).values[0],
                "minority": df['Board members (minority) as % of total'].tail(1).values[0]
            },
            {
                "category": "Management",
                "men": df['Employees in all management positions (male) % annual'].tail(1).values[0],
                "women": df['Employees in all management positions (female) % annual'].tail(1).values[0],
                "minority": df['Employees in all management positions (minority) % annual'].tail(1).values[0]
            }
        ]

        results["wellness_data"] = df[["Year", "Employees covered by collective bargaining Persons annual"]].to_dict(orient='records')
        
        lrcs = labor_rights_compliance_score(df)
        results["labor_rights_compliance_score"] = lrcs

        results["gender_diversity_board"] = [
            { "name": "Female", "value": df["Board members (female) as % of total"].tail(1).values[0], "color": "#EC4899", "description": "Women board members" },
            { "name": "Male", "value": df["Board members (male) as % of total"].tail(1).values[0], "color": "#3B82F6", "description": "Men board members" },
            { "name": "Minority", "value": df["Board members (minority) as % of total"].tail(1).values[0], "color": "#f6f03bff", "description": "Minority board members" },
        ]
       
        results["anti_corruption_training"] = math.floor(100 * (df["Employees Trained (Anti-Corruption)"].tail(1).values[0] / df["Total number of employees"].tail(1).values[0]))
        
        results["disability_representation"] = math.floor(100 * (df["Board members with disabilities"].tail(1).values[0] / df["Total number of employees"].tail(1).values[0]))
        
        results["education_diversity_data"] = [
                { "name": "Business/MBA", "value": int(df["Board education Business"].tail(1).values[0]), 
                 "color": "#3B82F6", "description": "Business/MBA background" },

                { "name": "Law", "value": int(df["Board education Law"].tail(1).values[0]), 
                 "color": "#EC4899", "description": "Legal background" },

                { "name": "Engineering/Tech", "value": int(df["Board education Engineering"].tail(1).values[0]),
                 "color": "#10B981", "description": "Engineering/Technology" },

                { "name": "Finance", "value": int(df["Board education Finance/Econ"].tail(1).values[0]),
                 "color": "#F59E0B", "description": "Finance background" },
                 
                { "name": "Other", "value": int(df["Board education Others"].tail(1).values[0]),
                 "color": "#6B7280", "description": "Other educational backgrounds" },
            ]
        
        results["age_group_composition"] = [
            { "name": "30-45", "value": int(df["Age-group composition 30-45"].tail(1).values[0]),
             "color": "#06B6D4", "description": "Ages 30-45" },
             
            { "name": "46-60", "value": int(df["Age-group composition 46-60"].tail(1).values[0]),
             "color": "#8B5CF6", "description": "Ages 46-60" },

            { "name": "61+", "value": int(df["Age-group composition 61+"].tail(1).values[0]),
             "color": "#F59E0B", "description": "Ages 61+" },
        ]

        ethnic_diversity_data_aze = df["Board ethnicity-AZE"].tail(1).values[0]
        results["ethnic_diversity_data"] = [
            { "name": "Azerbaijani", "value": ethnic_diversity_data_aze,
             "color": "#8B5CF6", "description": "Azerbaijani" },

            { "name": "Others", "value": round(100 - ethnic_diversity_data_aze, 1), 
             "color": "#F59E0B", "description": "Others" },
        ]


        shareholder_data_raw = df[["Year",
            "shareholder percentages (broad composition).Pension fund",
            "shareholder percentages (broad composition). Ataturk shares", 
            "shareholder percentages (broad composition). Free float",
        ]].to_dict(orient='records')
    
        shareholder_data = []
        for item in shareholder_data_raw:
            shareholder_data.append({
                "Year": item["Year"],
                "Pension fund": item["shareholder percentages (broad composition).Pension fund"] * 100,
                "Ataturk shares": item["shareholder percentages (broad composition). Ataturk shares"] * 100,
                "Free float": item["shareholder percentages (broad composition). Free float"] * 100,
            })
        
        # shareholder_data["shareholder percentages (broad composition).Pension fund"] = shareholder_data["shareholder percentages (broad composition).Pension fund"] * 100
        # shareholder_data["shareholder percentages (broad composition). Ataturk shares"] = shareholder_data["shareholder percentages (broad composition). Ataturk shares"] * 100
        # shareholder_data["shareholder percentages (broad composition). Free float"] = shareholder_data["shareholder percentages (broad composition). Free float"] * 100

        results["shareholder_rights_data"] = shareholder_data


        return results

    except Exception as e:
        return JSONResponse(status_code=500, content={"message": f"There was an error processing the file: {e}"})