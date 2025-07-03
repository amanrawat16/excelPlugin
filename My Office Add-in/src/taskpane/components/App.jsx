import * as React from "react";
import { useState } from "react";

// Rich lead data for meaningful AI analysis
const SAMPLE_DATA = [
  ["Name", "Email", "Company", "Industry", "Revenue", "Employee Count", "Contact History", "Response Rate", "Budget Range", "Timeline", "Pain Points", "Decision Maker", "Last Contact", "Lead Source", "Engagement Level"],
  ["Sarah Chen", "sarah.chen@innovatech.com", "Innovatech Solutions", "Technology", "$12M", "85", "3 meetings, 5 emails, 2 demos", "92%", "$100K-$150K", "Q1 2024", "Legacy system migration, scalability issues", "CTO", "2024-01-15", "LinkedIn", "High"],
  ["Michael Rodriguez", "m.rodriguez@healthplus.com", "HealthPlus Systems", "Healthcare", "$28M", "120", "2 meetings, 3 emails, 1 proposal", "78%", "$75K-$100K", "Q2 2024", "HIPAA compliance, patient data security", "CIO", "2024-01-10", "Trade Show", "Medium"],
  ["Jennifer Park", "jennifer@fintechpro.com", "FinTech Pro", "Finance", "$45M", "200", "4 meetings, 8 emails, 3 demos", "88%", "$150K-$200K", "Q1 2024", "Regulatory compliance, real-time processing", "VP Engineering", "2024-01-18", "Referral", "High"],
  ["David Thompson", "d.thompson@retailgiant.com", "Retail Giant Inc", "Retail", "$120M", "500", "1 meeting, 2 emails", "45%", "$50K-$75K", "Q3 2024", "Inventory management, customer analytics", "IT Director", "2024-01-05", "Cold Call", "Low"],
  ["Lisa Wang", "lisa.wang@edutech.com", "EduTech Innovations", "Education", "$8M", "60", "2 meetings, 4 emails, 1 workshop", "85%", "$60K-$90K", "Q2 2024", "Student engagement, remote learning tools", "CTO", "2024-01-12", "Website", "Medium"],
  ["Robert Johnson", "r.johnson@manufacturing.com", "Advanced Manufacturing", "Manufacturing", "$35M", "150", "3 meetings, 6 emails, 2 site visits", "82%", "$80K-$120K", "Q1 2024", "Supply chain optimization, quality control", "Operations Director", "2024-01-20", "Industry Conference", "High"],
  ["Amanda Foster", "a.foster@consulting.com", "Strategic Consulting", "Consulting", "$18M", "90", "1 meeting, 3 emails", "55%", "$40K-$60K", "Q4 2024", "Client management, reporting tools", "Managing Partner", "2024-01-08", "Email Campaign", "Low"],
  ["Carlos Mendez", "c.mendez@logistics.com", "Global Logistics", "Logistics", "$65M", "300", "2 meetings, 5 emails, 1 demo", "70%", "$90K-$130K", "Q2 2024", "Route optimization, real-time tracking", "VP Operations", "2024-01-14", "Partner Referral", "Medium"],
  ["Rachel Green", "r.green@marketing.com", "Digital Marketing Pro", "Marketing", "$15M", "75", "3 meetings, 7 emails, 2 proposals", "90%", "$70K-$100K", "Q1 2024", "Campaign automation, ROI tracking", "Marketing Director", "2024-01-16", "Social Media", "High"],
  ["James Wilson", "j.wilson@realestate.com", "Premier Real Estate", "Real Estate", "$85M", "250", "4 meetings, 10 emails, 3 demos", "95%", "$120K-$180K", "Q1 2024", "Property management, client portal", "Technology VP", "2024-01-19", "Industry Event", "High"]
];

const App = () => {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [results, setResults] = useState(null);

  /**
   * Insert sample lead data into the worksheet for demo/testing.
   */
  const addSampleData = async () => {
    setInfo("");
    setError("");
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRangeByIndexes(0, 0, SAMPLE_DATA.length, SAMPLE_DATA[0].length);
        range.values = SAMPLE_DATA;
        range.format.autofitColumns();
        await context.sync();
        setInfo("Sample lead data added! Select the data and click 'Score Leads' to analyze.");
      });
    } catch (err) {
      setError("Error adding sample data: " + err.message);
    }
  };

  /**
   * Test Office.js network connectivity
   */
  const testNetwork = async () => {
    setInfo("");
    setError("");
    try {
      console.log("=== Testing Network Connectivity ===");
      
      // Test 1: Simple fetch to a public API
      console.log("Test 1: Public API test");
      const publicResponse = await fetch('https://httpbin.org/json');
      console.log("Public API status:", publicResponse.status);
      
      // Test 2: Test our backend health endpoint
      console.log("Test 2: Backend health test");
      const healthResponse = await fetch('https://nasty-ears-decide.loca.lt/health', {
        headers: {
          'bypass-tunnel-reminder': 'true',
          'User-Agent': 'Excel-Addin/1.0'
        }
      });
      console.log("Backend health status:", healthResponse.status);
      const healthData = await healthResponse.json();
      console.log("Backend health data:", healthData);
      
      // Test 3: Test our scoring endpoint with minimal data
      console.log("Test 3: Backend scoring test");
      const testResponse = await fetch('https://nasty-ears-decide.loca.lt/score-leads', {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Cache-Control': 'no-cache',
          'bypass-tunnel-reminder': 'true',
          'User-Agent': 'Excel-Addin/1.0'
        },
        body: JSON.stringify({ 
          leads: [{"Name": "Test Lead", "Email": "test@example.com", "Company": "Test Corp"}] 
        }),
        mode: 'cors',
        credentials: 'omit'
      });
      console.log("Backend scoring status:", testResponse.status);
      
      if (testResponse.ok) {
        const testData = await testResponse.json();
        console.log("Backend scoring data:", testData);
        setInfo("✅ All network tests passed! Backend is accessible.");
      } else {
        const errorText = await testResponse.text();
        throw new Error(`Backend scoring failed: ${errorText}`);
      }
      
    } catch (err) {
      console.error("Network test error:", err);
      setError(`❌ Network test failed: ${err.message}`);
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
   * Score leads using the Python backend with Gemini AI integration.
   */
  const scoreLeads = async () => {
    setLoading(true);
    setError("");
    setInfo("");
    setResults(null);

    console.log("=== Starting Lead Scoring ===");

    try {
      await Excel.run(async (context) => {
        console.log("1. Excel.run context created");
        
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        console.log("2. Got active worksheet");
        
        // Get used range
        const usedRange = worksheet.getUsedRange();
        usedRange.load("address, rowCount, columnCount");
        await context.sync();
        
        console.log("3. Used range loaded:", usedRange.address, `${usedRange.rowCount}x${usedRange.columnCount}`);
        
        if (!usedRange.address || usedRange.address === "") {
          setError("❌ No Data Found: The worksheet is empty. Please click 'Insert Sample Data' to add lead data first.");
          setLoading(false);
          return;
        }

        // Load values from the range
        console.log("4. Loading values from range...");
        
        let values;
        try {
          usedRange.load("values");
          await context.sync();
          values = usedRange.values;
          console.log("5. Values loaded successfully:", values);
          
          if (!values || values.length === 0) {
            throw new Error("No values found in the selected range");
          }
          
        } catch (loadError) {
          console.error("Load error:", loadError);
          throw new Error(`❌ Data Loading Failed: Unable to read data from Excel. Please try selecting a smaller range or restart Excel. Error: ${loadError.message}`);
        }

        if (!values || values.length < 2) {
          setError("❌ Insufficient Data: Need at least 2 rows (headers + data). Please click 'Insert Sample Data' first.");
          setLoading(false);
          return;
        }

        // Prepare data for API
        const headers = values[0];
        const leads = values.slice(1).map((row, index) => {
          const lead = {};
          headers.forEach((header, colIndex) => {
            if (header) {
              lead[String(header).trim()] = row[colIndex] || "";
            }
          });
          return lead;
        }).filter(lead => Object.keys(lead).length > 0);
        
        console.log("6. Processed leads:", leads.length, "leads");
        
        if (leads.length === 0) {
          setError("❌ No valid lead data found. Please check your data format.");
          setLoading(false);
          return;
        }
        
        // Send to backend using proper Office.js network request
        console.log("7. Sending to backend...");
        console.log("API URL:", 'https://nasty-ears-decide.loca.lt/score-leads');
        console.log("Request payload:", JSON.stringify({ leads }, null, 2));
        
        // Use proper Office.js network request with error handling
        let response;
        try {
          response = await fetch('https://nasty-ears-decide.loca.lt/score-leads', {
            method: 'POST',
            headers: { 
              'Content-Type': 'application/json',
              'Accept': 'application/json',
              'Cache-Control': 'no-cache',
              'bypass-tunnel-reminder': 'true',
              'User-Agent': 'Excel-Addin/1.0'
            },
            body: JSON.stringify({ leads }),
            mode: 'cors',
            credentials: 'omit'
          });
          
          console.log("Response status:", response.status);
          console.log("Response headers:", Object.fromEntries(response.headers.entries()));
          
          if (!response.ok) {
            const errorText = await response.text();
            console.error("Backend error response:", errorText);
            throw new Error(`Backend Error (${response.status}): ${errorText}`);
          }
        } catch (fetchError) {
          console.error("Fetch error details:", fetchError);
          throw new Error(`Network Error: ${fetchError.message}`);
        }
        
        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Backend Error (${response.status}): ${errorText}`);
        }
        
        const result = await response.json();
        console.log("8. Backend response received:", result);
        
        setResults(result);
        setInfo(`✅ Successfully scored ${leads.length} leads with AI! Check the results below.`);

        // Highlight rows based on scores
        try {
          console.log("9. Starting row highlighting...");
          for (let i = 0; i < Math.min(result.scored_leads.length, 5); i++) {
            const scoredLead = result.scored_leads[i];
            const rowRange = usedRange.getRow(i + 1);
            
            if (scoredLead.score >= 80) {
              rowRange.format.fill.color = "#D4EDDA"; // Green for high quality
            } else if (scoredLead.score >= 60) {
              rowRange.format.fill.color = "#FFF3CD"; // Yellow for medium quality
            } else {
              rowRange.format.fill.color = "#F8D7DA"; // Red for low quality
            }
          }
          await context.sync();
          console.log("10. Row highlighting completed");
        } catch (highlightError) {
          console.log("Highlighting failed, but scoring completed:", highlightError);
          setInfo(`Scoring completed successfully! Highlighting failed: ${highlightError.message}`);
        }
      });
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error("Error message:", err.message);
      console.error("Error stack:", err.stack);
      
      if (err.message.includes("❌")) {
        setError(err.message);
      } else {
        setError(`❌ Unexpected Error: ${err.message}`);
      }
    }
    
    setLoading(false);
  };

  return (
    <div style={{ padding: 16, fontFamily: 'Segoe UI, Arial, sans-serif' }}>
      <h2>AI-Powered Lead Scoring</h2>
      <p>Analyze leads with Gemini AI integration</p>
      <p style={{ fontSize: '12px', color: '#666', marginBottom: 16 }}>
        💡 <strong>Tip:</strong> Click "Insert Sample Data" first, then click "Score Leads" to analyze with AI.
      </p>
      
      <div style={{ marginBottom: 16 }}>
        <button 
          onClick={addSampleData} 
          style={{ 
            padding: '10px 16px', 
            fontSize: 14, 
            marginRight: 8,
            backgroundColor: '#0078d4',
            color: 'white',
            border: 'none',
            borderRadius: 4,
            cursor: 'pointer'
          }}
        >
          Insert Sample Data
        </button>
        <button 
          onClick={scoreLeads} 
          disabled={loading}
          style={{ 
            padding: '10px 16px', 
            fontSize: 14, 
            marginRight: 8,
            backgroundColor: loading ? '#ccc' : '#107c10',
            color: 'white',
            border: 'none',
            borderRadius: 4,
            cursor: loading ? 'not-allowed' : 'pointer'
          }}
        >
          {loading ? "Scoring with AI..." : "Score Leads"}
        </button>
        <button 
          onClick={clearFormatting}
          style={{ 
            padding: '10px 16px', 
            fontSize: 14,
            backgroundColor: '#d83b01',
            color: 'white',
            border: 'none',
            borderRadius: 4,
            cursor: 'pointer',
            marginRight: 8
          }}
        >
          Clear Formatting
        </button>
        <button 
          onClick={testNetwork}
          style={{ 
            padding: '10px 16px', 
            fontSize: 14,
            backgroundColor: '#6f42c1',
            color: 'white',
            border: 'none',
            borderRadius: 4,
            cursor: 'pointer'
          }}
        >
          Test Network
        </button>
      </div>

      {info && (
        <div style={{ 
          color: 'green', 
          marginBottom: 12, 
          padding: '8px 12px',
          backgroundColor: '#d4edda',
          borderRadius: 4,
          border: '1px solid #c3e6cb'
        }}>
          {info}
        </div>
      )}

      {error && (
        <div style={{ 
          color: '#721c24', 
          marginBottom: 12, 
          padding: '8px 12px',
          backgroundColor: '#f8d7da',
          borderRadius: 4,
          border: '1px solid #f5c6cb'
        }}>
          {error}
        </div>
      )}

      {results && (
        <div style={{ marginTop: 20 }}>
          <h3>AI Lead Scoring Results</h3>
          
          <div style={{ marginBottom: 16 }}>
            <h4>Summary</h4>
            <p><strong>Total Leads:</strong> {results.summary.total_leads}</p>
            <p><strong>High Quality (80+):</strong> <span style={{color: 'green'}}>{results.summary.high_quality}</span></p>
            <p><strong>Medium Quality (60-79):</strong> <span style={{color: 'orange'}}>{results.summary.medium_quality}</span></p>
            <p><strong>Low Quality (&lt;60):</strong> <span style={{color: 'red'}}>{results.summary.low_quality}</span></p>
            <p><strong>Average Score:</strong> {results.summary.average_score.toFixed(1)}</p>
          </div>

          <div style={{ marginBottom: 16 }}>
            <h4>AI Insights</h4>
            <div style={{ 
              padding: '12px', 
              backgroundColor: '#f8f9fa', 
              borderRadius: 4,
              border: '1px solid #dee2e6'
            }}>
              {results.ai_insights}
            </div>
          </div>

          <div>
            <h4>Individual Lead Scores</h4>
            <div style={{ maxHeight: '300px', overflowY: 'auto' }}>
              {results.scored_leads.map((lead, index) => (
                <div key={index} style={{ 
                  marginBottom: 8, 
                  padding: '8px', 
                  backgroundColor: '#f8f9fa',
                  borderRadius: 4,
                  border: '1px solid #dee2e6'
                }}>
                  <div style={{ fontWeight: 'bold' }}>{lead.name} - {lead.company}</div>
                  <div>Score: <strong style={{color: lead.score >= 80 ? 'green' : lead.score >= 60 ? 'orange' : 'red'}}>{lead.score}</strong></div>
                  <div style={{ fontSize: '12px', color: '#666', marginTop: 4 }}>
                    {lead.reasoning}
                  </div>
                  <div style={{ fontSize: '12px', color: '#0078d4', marginTop: 4, fontStyle: 'italic' }}>
                    {lead.ai_insights}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
