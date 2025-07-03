from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Dict, Any
import google.generativeai as genai
import os
from dotenv import load_dotenv
import json

# Load environment variables
load_dotenv()

# Configure Gemini AI
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
try:
    model = genai.GenerativeModel('gemini-1.5-flash')
except:
    model = genai.GenerativeModel('gemini-pro')

app = FastAPI(title="Lead Scoring API", version="1.0.0")

# Enable CORS for Excel add-in with proper Office.js configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://localhost:3000",
        "https://breezy-fans-see.loca.lt",
        "https://nasty-ears-decide.loca.lt",
        "https://appsforoffice.microsoft.com",
        "https://secure.aadcdn.microsoftonline-p.com",
        "https://az689774.vo.msecnd.net",
        "https://nexus.officeapps.live.com",
        "https://browser.pipe.aria.microsoft.com",
        "https://telemetryservice.firstpartyapps.oaspapps.com"
    ],
    allow_credentials=False,  # Office.js doesn't send credentials
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["*"]
)

class LeadData(BaseModel):
    name: str = ""
    email: str = ""
    company: str = ""
    phone: str = ""
    industry: str = ""
    revenue: str = ""
    contact_history: str = ""
    response_rate: str = ""

class LeadAnalysisRequest(BaseModel):
    leads: List[Dict[str, Any]]

class LeadScore(BaseModel):
    name: str
    email: str
    company: str
    quality_score: int
    recommendation: str
    next_action: str
    risk_factors: List[str]
    ai_insights: str

def calculate_lead_score(lead_data: Dict[str, Any]) -> Dict[str, Any]:
    """Calculate lead score based on comprehensive business factors"""
    score = 0
    risk_factors = []
    strengths = []
    
    # Extract fields with defaults
    name = str(lead_data.get("Name", "")).strip()
    email = str(lead_data.get("Email", "")).strip()
    company = str(lead_data.get("Company", "")).strip()
    industry = str(lead_data.get("Industry", "")).strip()
    revenue = str(lead_data.get("Revenue", "")).strip()
    employee_count = str(lead_data.get("Employee Count", "")).strip()
    contact_history = str(lead_data.get("Contact History", "")).strip()
    response_rate = str(lead_data.get("Response Rate", "")).strip()
    budget_range = str(lead_data.get("Budget Range", "")).strip()
    timeline = str(lead_data.get("Timeline", "")).strip()
    pain_points = str(lead_data.get("Pain Points", "")).strip()
    decision_maker = str(lead_data.get("Decision Maker", "")).strip()
    lead_source = str(lead_data.get("Lead Source", "")).strip()
    engagement_level = str(lead_data.get("Engagement Level", "")).strip()
    
    # Basic validation (10 points)
    if name and email and company:
        score += 10
        strengths.append("Complete contact information")
    elif name and email:
        score += 7
        risk_factors.append("Missing company information")
    elif name and company:
        score += 5
        risk_factors.append("Missing email")
    else:
        score += 2
        risk_factors.append("Missing basic contact information")
    
    # Email validation (5 points)
    if "@" in email and "." in email:
        score += 5
    else:
        score += 0
        risk_factors.append("Invalid email format")
    
    # Industry analysis (10 points)
    if industry:
        if industry.lower() in ["technology", "healthcare", "finance"]:
            score += 10
            strengths.append("High-value industry")
        elif industry.lower() in ["manufacturing", "retail", "consulting"]:
            score += 8
            strengths.append("Good industry fit")
        else:
            score += 5
    else:
        score += 0
        risk_factors.append("Industry not specified")
    
    # Revenue analysis (15 points)
    if revenue:
        if "M" in revenue.upper() and any(char.isdigit() for char in revenue):
            revenue_num = float(''.join(filter(str.isdigit, revenue)))
            if revenue_num >= 50:
                score += 15
                strengths.append("Large enterprise")
            elif revenue_num >= 20:
                score += 12
                strengths.append("Mid-market company")
            elif revenue_num >= 10:
                score += 10
                strengths.append("Established business")
            else:
                score += 7
        else:
            score += 3
            risk_factors.append("Revenue information unclear")
    else:
        score += 0
        risk_factors.append("Revenue not specified")
    
    # Employee count analysis (5 points)
    if employee_count:
        try:
            emp_count = int(employee_count)
            if emp_count >= 200:
                score += 5
                strengths.append("Large team")
            elif emp_count >= 100:
                score += 4
                strengths.append("Growing company")
            elif emp_count >= 50:
                score += 3
            else:
                score += 2
        except:
            score += 1
    else:
        score += 0
        risk_factors.append("Employee count not specified")
    
    # Response rate analysis (10 points)
    if response_rate:
        try:
            rate = float(response_rate.replace("%", ""))
            if rate >= 90:
                score += 10
                strengths.append("Excellent responsiveness")
            elif rate >= 80:
                score += 8
                strengths.append("Good responsiveness")
            elif rate >= 60:
                score += 6
            elif rate >= 40:
                score += 4
                risk_factors.append("Low response rate")
            else:
                score += 2
                risk_factors.append("Very low response rate")
        except:
            score += 2
            risk_factors.append("Response rate unclear")
    else:
        score += 0
        risk_factors.append("Response rate not specified")
    
    # Budget analysis (15 points)
    if budget_range:
        if "150K" in budget_range or "200K" in budget_range:
            score += 15
            strengths.append("High budget")
        elif "100K" in budget_range:
            score += 12
            strengths.append("Good budget")
        elif "75K" in budget_range:
            score += 10
        elif "50K" in budget_range:
            score += 7
        else:
            score += 3
            risk_factors.append("Limited budget")
    else:
        score += 0
        risk_factors.append("Budget not specified")
    
    # Timeline analysis (10 points)
    if timeline:
        if "Q1" in timeline:
            score += 10
            strengths.append("Immediate timeline")
        elif "Q2" in timeline:
            score += 8
            strengths.append("Near-term timeline")
        elif "Q3" in timeline:
            score += 5
        elif "Q4" in timeline:
            score += 3
            risk_factors.append("Long timeline")
        else:
            score += 2
            risk_factors.append("Timeline unclear")
    else:
        score += 0
        risk_factors.append("Timeline not specified")
    
    # Decision maker analysis (10 points)
    if decision_maker:
        if any(title in decision_maker.upper() for title in ["CTO", "CIO", "VP", "DIRECTOR"]):
            score += 10
            strengths.append("Senior decision maker")
        elif any(title in decision_maker.upper() for title in ["MANAGER", "LEAD"]):
            score += 7
            strengths.append("Mid-level decision maker")
        else:
            score += 5
    else:
        score += 0
        risk_factors.append("Decision maker not specified")
    
    # Engagement level analysis (10 points)
    if engagement_level:
        if engagement_level.lower() == "high":
            score += 10
            strengths.append("High engagement")
        elif engagement_level.lower() == "medium":
            score += 7
            strengths.append("Moderate engagement")
        elif engagement_level.lower() == "low":
            score += 3
            risk_factors.append("Low engagement")
        else:
            score += 5
    else:
        score += 0
        risk_factors.append("Engagement level not specified")
    
    # Lead source analysis (5 points)
    if lead_source:
        if lead_source.lower() in ["referral", "industry conference", "trade show"]:
            score += 5
            strengths.append("Quality lead source")
        elif lead_source.lower() in ["linkedin", "website", "social media"]:
            score += 4
        elif lead_source.lower() in ["cold call", "email campaign"]:
            score += 2
        else:
            score += 3
    else:
        score += 0
        risk_factors.append("Lead source not specified")
    
    return {
        "score": min(score, 100),
        "risk_factors": risk_factors,
        "strengths": strengths
    }

def get_ai_insights(lead_data: Dict[str, Any], score_data: Dict[str, Any]) -> str:
    """Get AI insights using Gemini"""
    try:
        name = lead_data.get("Name", "")
        email = lead_data.get("Email", "")
        company = lead_data.get("Company", "")
        industry = lead_data.get("Industry", "")
        revenue = lead_data.get("Revenue", "")
        employee_count = lead_data.get("Employee Count", "")
        contact_history = lead_data.get("Contact History", "")
        response_rate = lead_data.get("Response Rate", "")
        budget_range = lead_data.get("Budget Range", "")
        timeline = lead_data.get("Timeline", "")
        pain_points = lead_data.get("Pain Points", "")
        decision_maker = lead_data.get("Decision Maker", "")
        lead_source = lead_data.get("Lead Source", "")
        engagement_level = lead_data.get("Engagement Level", "")
        
        prompt = f"""
        Analyze this sales lead and provide strategic insights:
        
        LEAD INFORMATION:
        - Name: {name} ({email})
        - Company: {company} ({industry} industry)
        - Revenue: {revenue} | Employees: {employee_count}
        - Decision Maker: {decision_maker}
        
        ENGAGEMENT METRICS:
        - Contact History: {contact_history}
        - Response Rate: {response_rate}
        - Engagement Level: {engagement_level}
        - Lead Source: {lead_source}
        
        BUSINESS CONTEXT:
        - Budget Range: {budget_range}
        - Timeline: {timeline}
        - Pain Points: {pain_points}
        
        QUALITY ASSESSMENT:
        - Score: {score_data['score']}/100
        - Strengths: {', '.join(score_data['strengths']) if score_data['strengths'] else 'None'}
        - Risk Factors: {', '.join(score_data['risk_factors']) if score_data['risk_factors'] else 'None'}
        
        Provide a comprehensive analysis (3-4 sentences) covering:
        1. Key business opportunities and strengths
        2. Potential challenges or concerns
        3. Strategic recommendations for next steps
        4. Priority level and urgency
        """
        
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        # Enhanced fallback analysis without AI
        score = score_data['score']
        strengths = score_data.get('strengths', [])
        risk_factors = score_data.get('risk_factors', [])
        
        name = lead_data.get("Name", "")
        company = lead_data.get("Company", "")
        industry = lead_data.get("Industry", "")
        budget_range = lead_data.get("Budget Range", "")
        timeline = lead_data.get("Timeline", "")
        decision_maker = lead_data.get("Decision Maker", "")
        
        if score >= 85:
            return f"EXCELLENT LEAD ({score}/100): {name} from {company} ({industry}) represents a high-priority opportunity. Key strengths: {', '.join(strengths[:3])}. With budget {budget_range} and timeline {timeline}, this {decision_maker} should be contacted immediately for proposal presentation."
        elif score >= 70:
            return f"STRONG LEAD ({score}/100): {name} from {company} shows good potential with {', '.join(strengths[:2])}. Timeline {timeline} and budget {budget_range} indicate readiness. Focus on addressing: {', '.join(risk_factors[:2]) if risk_factors else 'No major concerns'}."
        elif score >= 50:
            return f"MODERATE LEAD ({score}/100): {name} from {company} needs nurturing. Strengths: {', '.join(strengths[:2]) if strengths else 'Limited'}. Concerns: {', '.join(risk_factors[:3]) if risk_factors else 'Missing information'}. Consider qualification call."
        else:
            return f"LOW-PRIORITY LEAD ({score}/100): {name} from {company} requires significant qualification. Major issues: {', '.join(risk_factors[:3]) if risk_factors else 'Multiple data gaps'}. Consider automated nurturing or archive."

def get_recommendation(score: int) -> str:
    """Get recommendation based on score"""
    if score >= 80:
        return "Hot Lead - High priority"
    elif score >= 60:
        return "Warm Lead - Follow up soon"
    elif score >= 40:
        return "Cold Lead - Nurture campaign"
    else:
        return "Dead Lead - Archive"

def get_next_action(score: int, lead: LeadData) -> str:
    """Get next action based on score and lead data"""
    if score >= 80:
        return "Schedule demo within 24 hours"
    elif score >= 60:
        return "Send proposal within 48 hours"
    elif score >= 40:
        return "Add to nurture campaign"
    else:
        return "Archive and focus on better leads"

@app.get("/")
async def root():
    return {"message": "Lead Scoring API is running"}

@app.post("/score-leads")
async def score_leads(request: LeadAnalysisRequest):
    """Score leads and return analysis with AI insights"""
    print("=== BACKEND: Received API call ===")
    print(f"Number of leads received: {len(request.leads)}")
    print(f"First lead data: {request.leads[0] if request.leads else 'No leads'}")
    print("=== END BACKEND LOG ===")
    
    results = []
    total_score = 0
    
    for lead_data in request.leads:
        # Calculate score
        score_data = calculate_lead_score(lead_data)
        total_score += score_data["score"]
        
        # Get AI insights
        ai_insights = get_ai_insights(lead_data, score_data)
        
        # Create result
        result = {
            "name": lead_data.get("Name", ""),
            "company": lead_data.get("Company", ""),
            "score": score_data["score"],
            "reasoning": f"Strengths: {', '.join(score_data['strengths'][:3]) if score_data['strengths'] else 'None'}. Issues: {', '.join(score_data['risk_factors'][:2]) if score_data['risk_factors'] else 'None'}.",
            "ai_insights": ai_insights
        }
        results.append(result)
    
    # Calculate summary
    avg_score = total_score / len(request.leads) if request.leads else 0
    high_quality = sum(1 for r in results if r["score"] >= 80)
    medium_quality = sum(1 for r in results if 60 <= r["score"] < 80)
    low_quality = sum(1 for r in results if r["score"] < 60)
    
    # Generate overall AI insights
    overall_insights = f"Analysis of {len(request.leads)} leads shows an average score of {avg_score:.1f}. "
    overall_insights += f"High-quality leads: {high_quality}, Medium: {medium_quality}, Low: {low_quality}. "
    if high_quality > len(request.leads) / 2:
        overall_insights += "Strong lead quality overall."
    elif low_quality > len(request.leads) / 2:
        overall_insights += "Focus on lead nurturing and qualification."
    else:
        overall_insights += "Mixed lead quality - prioritize high-scoring leads."
    
    return {
        "scored_leads": results,
        "summary": {
            "total_leads": len(request.leads),
            "average_score": avg_score,
            "high_quality": high_quality,
            "medium_quality": medium_quality,
            "low_quality": low_quality
        },
        "ai_insights": overall_insights
    }

@app.get("/health")
async def health_check():
    return {"status": "healthy", "service": "Lead Scoring API"}

@app.post("/test-data")
async def test_data(request: dict):
    """Test endpoint to see what data is being sent"""
    print("=== TEST ENDPOINT CALLED ===")
    print(f"Received data: {request}")
    print("=== END TEST LOG ===")
    return {"message": "Data received", "data": request}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 