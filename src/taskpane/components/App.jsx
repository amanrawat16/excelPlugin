import * as React from "react";
import { useState } from "react";

// Realistic lead data for demo
const SAMPLE_LEAD_DATA = [
  ["Name", "Email", "Company", "Contact Date", "Response", "Follow-up Date", "Notes", "Budget", "Timeline"],
  ["John Smith", "john@techcorp.com", "TechCorp Inc", "2024-01-15", "Interested in demo", "2024-01-20", "Asked for pricing, 50-100 employees", "$50K-100K", "Q1 2024"],
  ["Sarah Johnson", "sarah@startup.com", "StartupXYZ", "2024-01-10", "No response", "2024-01-17", "Sent 2 emails, no reply", "Unknown", "Unknown"],
  ["Mike Chen", "mike@enterprise.com", "Enterprise Solutions", "2024-01-12", "Not interested", "2024-01-18", "Budget constraints, maybe next year", "$0", "Q4 2024"],
  ["Lisa Brown", "lisa@growth.com", "GrowthCo", "2024-01-14", "Very interested", "2024-01-21", "Wants to see ROI case studies", "$100K+", "Q1 2024"],
  ["David Wilson", "david@small.com", "SmallBiz", "2024-01-16", "Need more info", "2024-01-22", "Asked about implementation time", "$10K-25K", "Q2 2024"],
  ["Emma Davis", "emma@consulting.com", "Consulting Partners", "2024-01-13", "Interested", "2024-01-19", "Decision maker, wants technical demo", "$75K-150K", "Q1 2024"],
  ["Alex Rodriguez", "alex@old.com", "Legacy Systems", "2024-01-11", "No response", "2024-01-16", "Sent 3 emails, called twice", "Unknown", "Unknown"],
  ["Rachel Green", "rachel@new.com", "New Ventures", "2024-01-15", "Love the concept", "2024-01-23", "Wants to pilot with 5 users", "$25K-50K", "Q1 2024"]
];

const App = () => {
  const [summary, setSummary] = useState(null);
  const [leadScores, setLeadScores] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  /**
   * Insert sample lead data into the worksheet for demo/testing.
   */
  const addSampleData = async () => {
    setInfo("");
    setError("");
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRangeByIndexes(0, 0, SAMPLE_LEAD_DATA.length, SAMPLE_LEAD_DATA[0].length);
        range.values = SAMPLE_LEAD_DATA;
        range.format.autofitColumns();
        await context.sync();
        setInfo("Sample lead data added to A1:I9.");
      });
    } catch (err) {
      setError("Error adding sample data: " + err.message);
    }
  };

  /**
   * Clear all background color formatting from the selected range.
   */
  const clearFormatting = async () => {
    setInfo("");
    setError("");
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.clear();
        await context.sync();
        setInfo("Formatting cleared for selected range.");
      });
    } catch (err) {
      setError("Error clearing formatting: " + err.message);
    }
  };

  /**
   * Analyze leads using the Python FastAPI backend with Gemini AI.
   */
  const analyzeLeads = async () => {
    setLoading(true);
    setError("");
    setSummary(null);
    setLeadScores([]);
    
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
        await context.sync();
        
        const values = range.values;
        if (!values || values.length < 2) {
          setError("Please select a table with headers and at least one row of lead data.");
          setLoading(false);
          return;
        }

        // Parse headers and data
        const headers = values[0];
        const headerMap = {};
        headers.forEach((h, i) => { headerMap[String(h).trim()] = i; });

        // Prepare leads data for API
        const leads = [];
        for (let i = 1; i < values.length; i++) {
          const row = values[i];
          const lead = {
            name: row[headerMap["Name"]] || "",
            email: row[headerMap["Email"]] || "",
            company: row[headerMap["Company"]] || "",
            contact_date: row[headerMap["Contact Date"]] || "",
            response: row[headerMap["Response"]] || "",
            follow_up_date: row[headerMap["Follow-up Date"]] || "",
            notes: row[headerMap["Notes"]] || "",
            budget: row[headerMap["Budget"]] || "",
            timeline: row[headerMap["Timeline"]] || ""
          };
          leads.push(lead);
        }

        // Call Python FastAPI backend
        const response = await fetch("http://localhost:8000/analyze-leads", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ leads }),
        });

        if (!response.ok) {
          throw new Error(`API Error: ${response.status} ${response.statusText}`);
        }

        const results = await response.json();
        setLeadScores(results);

        // Highlight rows based on scores
        for (let i = 1; i < values.length; i++) {
          const rowRange = range.getRow(i);
          const score = results[i - 1]?.quality_score || 0;
          
          if (score >= 80) {
            rowRange.format.fill.color = "#D4EDDA"; // Green for hot leads
          } else if (score >= 60) {
            rowRange.format.fill.color = "#FFF3CD"; // Yellow for warm leads
          } else if (score >= 40) {
            rowRange.format.fill.color = "#F8D7DA"; // Light red for cold leads
          } else {
            rowRange.format.fill.color = "#F5C6CB"; // Red for dead leads
          }
        }

        await context.sync();

        // Calculate summary
        const total = results.length;
        const hotLeads = results.filter(r => r.quality_score >= 80).length;
        const warmLeads = results.filter(r => r.quality_score >= 60 && r.quality_score < 80).length;
        const coldLeads = results.filter(r => r.quality_score >= 40 && r.quality_score < 60).length;
        const deadLeads = results.filter(r => r.quality_score < 40).length;

        setSummary({
          total,
          hotLeads,
          warmLeads,
          coldLeads,
          deadLeads,
          averageScore: Math.round(results.reduce((sum, r) => sum + r.quality_score, 0) / total)
        });

      });
    } catch (err) {
      setError("Error analyzing leads: " + err.message);
      console.error("Analysis error:", err);
    }
    
    setLoading(false);
  };

  return (
    <div style={{ padding: 16, fontFamily: 'Segoe UI, Arial, sans-serif' }}>
      <h2>Lead Scoring & Analysis</h2>
      <p>Select a table of leads and get AI-powered scoring and insights:</p>
      
      <button onClick={addSampleData} style={{ padding: '8px 16px', fontSize: 16, marginBottom: 12 }}>
        Add Sample Lead Data
      </button>
      <br />
      
      <button onClick={analyzeLeads} disabled={loading} style={{ padding: '8px 16px', fontSize: 16, marginRight: 8 }}>
        {loading ? "Analyzing..." : "Analyze Leads"}
      </button>
      
      <button onClick={clearFormatting} style={{ padding: '8px 16px', fontSize: 16 }}>
        Clear Formatting
      </button>
      
      {info && <div style={{ color: 'green', marginTop: 12 }}>{info}</div>}
      {error && <div style={{ color: 'red', marginTop: 12 }}>{error}</div>}
      
      {summary && (
        <div style={{ marginTop: 20 }}>
          <h3>Lead Analysis Summary</h3>
          <p>Total leads analyzed: <b>{summary.total}</b></p>
          <p>Average score: <b>{summary.averageScore}/100</b></p>
          <p>Hot leads (80+): <b style={{ color: 'green' }}>{summary.hotLeads}</b></p>
          <p>Warm leads (60-79): <b style={{ color: 'orange' }}>{summary.warmLeads}</b></p>
          <p>Cold leads (40-59): <b style={{ color: '#dc3545' }}>{summary.coldLeads}</b></p>
          <p>Dead leads (<40): <b style={{ color: '#721c24' }}>{summary.deadLeads}</b></p>
        </div>
      )}
      
      {leadScores.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3>Individual Lead Scores</h3>
          <div style={{ maxHeight: '300px', overflowY: 'auto' }}>
            {leadScores.map((lead, idx) => (
              <div key={idx} style={{ 
                border: '1px solid #ddd', 
                margin: '8px 0', 
                padding: '12px', 
                borderRadius: '4px',
                backgroundColor: lead.quality_score >= 80 ? '#f8fff8' : '#fff8f8'
              }}>
                <h4 style={{ margin: '0 0 8px 0' }}>
                  {lead.name} - {lead.company}
                </h4>
                <p style={{ margin: '4px 0' }}>
                  <strong>Score:</strong> {lead.quality_score}/100 - {lead.recommendation}
                </p>
                <p style={{ margin: '4px 0' }}>
                  <strong>Next Action:</strong> {lead.next_action}
                </p>
                {lead.risk_factors.length > 0 && (
                  <p style={{ margin: '4px 0', color: '#dc3545' }}>
                    <strong>Risk Factors:</strong> {lead.risk_factors.join(', ')}
                  </p>
                )}
                <p style={{ margin: '4px 0', fontSize: '12px', fontStyle: 'italic' }}>
                  <strong>AI Insights:</strong> {lead.ai_insights}
                </p>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

export default App; 