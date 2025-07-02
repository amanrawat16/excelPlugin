import * as React from "react";
import { useState } from "react";

// Required fields for validation
const REQUIRED_FIELDS = ["Name", "Email", "Phone", "Company"];

// Sample data for demo/testing
const SAMPLE_DATA = [
  ["Name", "Email", "Phone", "Company"],
  ["Alice Smith", "alice@example.com", "1234567890", "Acme Corp"],
  ["Bob Jones", "", "9876543210", "Beta LLC"], // Missing email
  ["Carol Lee", "carol@company.com", "", "Gamma Inc"], // Missing phone
  ["", "dave@domain.com", "1234567890", "Delta Ltd"], // Missing name
  ["Eve Adams", "eve@adams.com", "12345", ""], // Short phone, missing company
  ["Frank", "frank[at]mail.com", "1234567890", "Zeta Org"], // Invalid email format
];

/**
 * Validate a single row of contact data.
 * @param {Array} row - The row data as an array of cell values.
 * @param {Object} headerMap - Maps field names to column indices.
 * @returns {Array} - List of missing or invalid fields for this row.
 */
const validateRow = (row, headerMap) => {
  const missing = [];
  REQUIRED_FIELDS.forEach((field) => {
    const idx = headerMap[field];
    if (idx === undefined || !row[idx] || String(row[idx]).trim() === "") {
      missing.push(field);
    } else if (field === "Email" && !String(row[idx]).includes("@")) {
      missing.push("Email (invalid format)");
    } else if (field === "Phone" && String(row[idx]).replace(/\D/g, "").length < 10) {
      missing.push("Phone (too short)");
    }
  });
  return missing;
};

const App = () => {
  // State for summary, invalid rows, loading, error, and info messages
  const [summary, setSummary] = useState(null);
  const [invalidRows, setInvalidRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  /**
   * Insert sample contact data into the worksheet for demo/testing.
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
        setInfo("Sample data added to A1:D7.");
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
   * Validate all rows in the selected range and highlight valid/invalid rows.
   */
  const checkContacts = async () => {
    setLoading(true);
    setError("");
    setSummary(null);
    setInvalidRows([]);
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
        await context.sync();
        const values = range.values;
        if (!values || values.length < 2) {
          setError("Please select a table with headers and at least one row of data.");
          setLoading(false);
          return;
        }
        // Map header names to column indices
        const headers = values[0];
        const headerMap = {};
        headers.forEach((h, i) => { headerMap[String(h).trim()] = i; });
        const invalid = [];
        for (let i = 1; i < values.length; i++) {
          const row = values[i];
          const missing = validateRow(row, headerMap);
          const rowRange = range.getRow(i);
          if (missing.length > 0) {
            invalid.push({ row: i + range.rowIndex, missing });
            // Highlight the invalid row in yellow
            rowRange.format.fill.color = "#FFF3CD";
          } else {
            // Highlight valid row in green
            rowRange.format.fill.color = "#D4EDDA";
          }
        }
        await context.sync();
        setSummary({
          total: values.length - 1,
          invalid: invalid.length,
          valid: values.length - 1 - invalid.length,
        });
        setInvalidRows(invalid);
      });
    } catch (err) {
      setError("Error: " + err.message);
    }
    setLoading(false);
  };

  return (
    <div style={{ padding: 16, fontFamily: 'Segoe UI, Arial, sans-serif' }}>
      <h2>Invalid Contact Checker</h2>
      <p>Select a table of leads (with columns: Name, Email, Phone, Company) and click below:</p>
      <button onClick={addSampleData} style={{ padding: '8px 16px', fontSize: 16, marginBottom: 12 }}>
        Add Sample Data
      </button>
      <br />
      <button onClick={checkContacts} disabled={loading} style={{ padding: '8px 16px', fontSize: 16, marginRight: 8 }}>
        {loading ? "Checking..." : "Check Contacts"}
      </button>
      <button onClick={clearFormatting} style={{ padding: '8px 16px', fontSize: 16 }}>
        Clear Formatting
      </button>
      {info && <div style={{ color: 'green', marginTop: 12 }}>{info}</div>}
      {error && <div style={{ color: 'red', marginTop: 12 }}>{error}</div>}
      {summary && (
        <div style={{ marginTop: 20 }}>
          <h3>Summary</h3>
          <p>Total leads checked: <b>{summary.total}</b></p>
          <p>Valid leads: <b style={{ color: 'green' }}>{summary.valid}</b></p>
          <p>Invalid leads: <b style={{ color: 'orange' }}>{summary.invalid}</b></p>
          {invalidRows.length > 0 && (
            <div style={{ marginTop: 10 }}>
              <h4>Invalid Rows</h4>
              <ul>
                {invalidRows.map((row, idx) => (
                  <li key={idx}>
                    Row {row.row + 1}: Missing/Invalid - {row.missing.join(", ")}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default App;
