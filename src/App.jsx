import { useState } from "react";
import "./App.css";

function App() {
  const [file, setFile] = useState(null);

  // Product information
  const [itemCode, setItemCode] = useState("");
  const [category, setCategory] = useState("");
  const [hsn, setHsn] = useState("");
  const [itemLocation, setItemLocation] = useState("");
  const [minimumStock, setMinimumStock] = useState("3");

  // Pricing configuration
  const [transport, setTransport] = useState("");
  const [gst, setGst] = useState("");
  const [extraCharge, setExtraCharge] = useState("");
  const [profit, setProfit] = useState("");

  const [loading, setLoading] = useState(false);
  const [dragActive, setDragActive] = useState(false);

  const validateFile = (selectedFile) => {
    if (!selectedFile) return false;

    if (
      selectedFile.type !== "application/pdf" &&
      !selectedFile.name.toLowerCase().endsWith(".pdf")
    ) {
      alert("Please upload a PDF file only.");
      return false;
    }

    if (selectedFile.size > 10 * 1024 * 1024) {
      alert("PDF size should be less than 10 MB.");
      return false;
    }

    setFile(selectedFile);
    return true;
  };

  const handleFileChange = (event) => {
    const selectedFile = event.target.files[0];

    if (selectedFile) {
      validateFile(selectedFile);
    }

    event.target.value = "";
  };

  const handleDrop = (event) => {
    event.preventDefault();
    setDragActive(false);

    const droppedFile = event.dataTransfer.files[0];

    if (droppedFile) {
      validateFile(droppedFile);
    }
  };

  const handleDragOver = (event) => {
    event.preventDefault();
    setDragActive(true);
  };

  const handleDragLeave = () => {
    setDragActive(false);
  };

  const removeFile = () => {
    setFile(null);
  };

  const handleGenerate = async () => {
    if (!file) {
      alert("Please upload a PDF first.");
      return;
    }

    if (!itemCode.trim()) {
      alert("Please enter Item Code.");
      return;
    }

    if (!category.trim()) {
      alert("Please enter Category.");
      return;
    }

    if (!hsn.trim()) {
      alert("Please enter HSN.");
      return;
    }

    if (!itemLocation.trim()) {
      alert("Please enter Item Location.");
      return;
    }

    if (minimumStock === "") {
      alert("Please enter Minimum Stock Quantity.");
      return;
    }

    if (transport === "") {
      alert("Please enter Transport Charge.");
      return;
    }

    if (gst === "") {
      alert("Please enter GST percentage. Enter 0 if GST is not applicable.");
      return;
    }

    if (extraCharge === "") {
      alert("Please enter Any Extra Charge. Enter 0 if none.");
      return;
    }

    if (profit === "") {
      alert("Please enter Profit Percentage.");
      return;
    }

    try {
      setLoading(true);

      const formData = new FormData();

      // PDF
      formData.append("pdf", file);

      // Product information
      formData.append("item_code", itemCode);
      formData.append("category", category);
      formData.append("hsn", hsn);
      formData.append("item_location", itemLocation);
      formData.append("minimum_stock", minimumStock);

      // Pricing information
      formData.append("transport_charge", transport);

      // GST is entered as percentage
      formData.append("gst_rate", gst);

      // Backend currently accepts GST amount also.
      // GST amount will be calculated by backend from GST rate.
      formData.append("gst_amount", "0");

      formData.append("any_extra_charge", extraCharge);
      formData.append("profit_percentage", profit);

      const response = await fetch(
        `${import.meta.env.VITE_API_URL || "http://127.0.0.1:5000"}/api/generate-excel`,
        {
          method: "POST",
          body: formData,
        }
      );

      if (!response.ok) {
        const errorData = await response.json().catch(() => null);

        throw new Error(
          errorData?.error ||
            errorData?.message ||
            "Excel generation failed."
        );
      }

      const blob = await response.blob();

      const downloadUrl = window.URL.createObjectURL(blob);

      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = "product_import.xlsx";

      document.body.appendChild(link);
      link.click();
      link.remove();

      window.URL.revokeObjectURL(downloadUrl);

      alert("Excel file generated successfully!");
    } catch (error) {
      console.error(error);
      alert(error.message || "Something went wrong.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app">

      {/* =========================
          HEADER / NAVBAR
      ========================= */}

      <header className="topbar">
        <div className="topbar-inner">

          <div className="brand">
            <div className="brand-icon">
              ⇄
            </div>

            <div>
              <div className="brand-name">
                PDF2Excel
              </div>

              <div className="brand-subtitle">
                Product Import Tool
              </div>
            </div>
          </div>

          <div className="status-badge">
            <span className="status-dot"></span>
            System Ready
          </div>

        </div>
      </header>


      {/* =========================
          MAIN DASHBOARD
      ========================= */}

      <main className="dashboard">

        {/* HERO */}

        <section className="hero">

          <div className="hero-content">

            <div className="hero-badge">
              <span>⚡</span>
              Automated Product Data Processing
            </div>

            <h1>
              Convert PDF Invoices
              <br />
              <span>Into Excel Import Files</span>
            </h1>

            <p>
              Upload your product invoice PDF and automatically
              extract product details, calculate selling prices,
              and generate a standardized Excel file.
            </p>

          </div>

          <div className="hero-visual">

            <div className="visual-card pdf-card">
              <div className="visual-icon">📄</div>
              <span>PDF</span>
            </div>

            <div className="visual-arrow">
              →
            </div>

            <div className="visual-card excel-card">
              <div className="visual-icon">📊</div>
              <span>Excel</span>
            </div>

          </div>

        </section>


        {/* PROCESS STEPS */}

        <section className="steps">

          <div className="step active">
            <div className="step-number">
              01
            </div>

            <div>
              <strong>Upload PDF</strong>
              <span>Select your invoice</span>
            </div>
          </div>

          <div className="step-line"></div>

          <div className="step">
            <div className="step-number">
              02
            </div>

            <div>
              <strong>Process Data</strong>
              <span>Extract & calculate</span>
            </div>
          </div>

          <div className="step-line"></div>

          <div className="step">
            <div className="step-number">
              03
            </div>

            <div>
              <strong>Download Excel</strong>
              <span>Ready to import</span>
            </div>
          </div>

        </section>


        {/* MAIN GRID */}

        <div className="dashboard-grid">

          {/* =========================
              LEFT SIDE
          ========================= */}

          <section className="main-card">

            <div className="card-heading">

              <div>
                <div className="section-label">
                  STEP 01
                </div>

                <h2>
                  Upload Product PDF
                </h2>

                <p>
                  Upload the invoice containing your
                  product information.
                </p>
              </div>

              <div className="card-icon">
                📄
              </div>

            </div>


            {/* UPLOAD AREA */}

            <label
              className={`upload-zone ${
                dragActive ? "drag-active" : ""
              } ${file ? "has-file" : ""}`}
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
            >

              <input
                type="file"
                accept=".pdf,application/pdf"
                onChange={handleFileChange}
              />

              {!file ? (
                <div className="upload-empty">

                  <div className="upload-circle">
                    ↑
                  </div>

                  <h3>
                    Drag & drop your PDF here
                  </h3>

                  <p>
                    or click to browse from your computer
                  </p>

                  <div className="upload-limit">
                    <span>PDF</span>
                    <span>•</span>
                    <span>Maximum 10 MB</span>
                  </div>

                </div>
              ) : (
                <div className="selected-file">

                  <div className="file-icon">
                    PDF
                  </div>

                  <div className="file-details">

                    <h3>
                      {file.name}
                    </h3>

                    <p>
                      {(file.size / 1024 / 1024).toFixed(2)} MB
                      &nbsp; • &nbsp;
                      PDF Document
                    </p>

                  </div>

                  <button
                    type="button"
                    className="remove-file"
                    onClick={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      removeFile();
                    }}
                  >
                    Remove
                  </button>

                </div>
              )}

            </label>


            {/* FILE INFO */}

            <div className="upload-note">
              <span>🔒</span>
              Your PDF is processed locally through your
              configured backend and is not permanently stored
              by this interface.
            </div>

          </section>


          {/* =========================
              RIGHT SIDE
          ========================= */}

          <section className="main-card pricing-card">

            <div className="card-heading">

              <div>
                <div className="section-label">
                  STEP 02
                </div>

                <h2>
                  Product & Pricing Configuration
                </h2>

                <p>
                  Enter product details and values used for price calculation.
                </p>
              </div>

              <div className="card-icon blue">
                ⚙
              </div>

            </div>


            <div className="pricing-fields">

              {/* ITEM CODE */}

              <div className="pricing-field">

                <label htmlFor="itemCode">
                  <span className="field-icon">
                    🏷️
                  </span>

                  Item Code
                </label>

                <div className="modern-input">

                  <input
                    id="itemCode"
                    type="text"
                    value={itemCode}
                    onChange={(event) =>
                      setItemCode(event.target.value)
                    }
                    placeholder="e.g. MG1"
                  />

                </div>

                <small>
                  Item code used in the Excel import file
                </small>

              </div>


              {/* CATEGORY */}

              <div className="pricing-field">

                <label htmlFor="category">
                  <span className="field-icon">
                    📦
                  </span>

                  Category
                </label>

                <div className="modern-input">

                  <input
                    id="category"
                    type="text"
                    value={category}
                    onChange={(event) =>
                      setCategory(event.target.value)
                    }
                    placeholder="e.g. GIFT"
                  />

                </div>

                <small>
                  Product category
                </small>

              </div>


              {/* HSN */}

              <div className="pricing-field">

                <label htmlFor="hsn">
                  <span className="field-icon">
                    🔢
                  </span>

                  HSN
                </label>

                <div className="modern-input">

                  <input
                    id="hsn"
                    type="text"
                    value={hsn}
                    onChange={(event) =>
                      setHsn(event.target.value)
                    }
                    placeholder="e.g. 12345"
                  />

                </div>

                <small>
                  HSN code for the products
                </small>

              </div>


              {/* ITEM LOCATION */}

              <div className="pricing-field">

                <label htmlFor="itemLocation">
                  <span className="field-icon">
                    📍
                  </span>

                  Item Location
                </label>

                <div className="modern-input">

                  <input
                    id="itemLocation"
                    type="text"
                    value={itemLocation}
                    onChange={(event) =>
                      setItemLocation(event.target.value)
                    }
                    placeholder="e.g. Store 1"
                  />

                </div>

                <small>
                  Location where the item will be stored
                </small>

              </div>


              {/* MINIMUM STOCK */}

              <div className="pricing-field">

                <label htmlFor="minimumStock">
                  <span className="field-icon">
                    📊
                  </span>

                  Minimum Stock Quantity
                </label>

                <div className="modern-input">

                  <input
                    id="minimumStock"
                    type="number"
                    min="0"
                    value={minimumStock}
                    onChange={(event) =>
                      setMinimumStock(event.target.value)
                    }
                  />

                </div>

                <small>
                  Minimum quantity allowed before restocking
                </small>

              </div>


              {/* TRANSPORT */}

              <div className="pricing-field">

                <label htmlFor="transport">
                  <span className="field-icon">
                    🚚
                  </span>

                  Transport Charge
                </label>

                <div className="modern-input">

                  <span className="input-prefix">
                    ₹
                  </span>

                  <input
                    id="transport"
                    type="number"
                    min="0"
                    value={transport}
                    onChange={(event) =>
                      setTransport(event.target.value)
                    }
                    placeholder="0"
                  />

                </div>

                <small>
                  Transportation charge for the invoice
                </small>

              </div>


              {/* GST */}

              <div className="pricing-field">

                <label htmlFor="gst">
                  <span className="field-icon">
                    🧾
                  </span>

                  GST
                </label>

                <div className="modern-input">

                  <input
                    id="gst"
                    type="number"
                    min="0"
                    value={gst}
                    onChange={(event) =>
                      setGst(event.target.value)
                    }
                    placeholder="0"
                  />

                  <span className="input-suffix">
                    %
                  </span>

                </div>

                <small>
                  Enter GST percentage. Use 0 if GST is not applicable.
                </small>

              </div>


              {/* EXTRA CHARGE */}

              <div className="pricing-field">

                <label htmlFor="extraCharge">
                  <span className="field-icon">
                    ➕
                  </span>

                  Any Extra Charge
                </label>

                <div className="modern-input">

                  <span className="input-prefix">
                    ₹
                  </span>

                  <input
                    id="extraCharge"
                    type="number"
                    min="0"
                    value={extraCharge}
                    onChange={(event) =>
                      setExtraCharge(event.target.value)
                    }
                    placeholder="0"
                  />

                </div>

                <small>
                  Enter any additional charge. Use 0 if none.
                </small>

              </div>


              {/* PROFIT */}

              <div className="pricing-field">

                <label htmlFor="profit">
                  <span className="field-icon">
                    📈
                  </span>

                  Profit Percentage
                </label>

                <div className="modern-input">

                  <input
                    id="profit"
                    type="number"
                    min="0"
                    value={profit}
                    onChange={(event) =>
                      setProfit(event.target.value)
                    }
                    placeholder="20"
                  />

                  <span className="input-suffix">
                    %
                  </span>

                </div>

                <small>
                  Profit percentage applied after extra charges
                </small>

              </div>

            </div>


            {/* CALCULATION PREVIEW */}

            <div className="calculation-box">

              <div className="calculation-title">
                Calculation Formula
              </div>

              <div className="formula">
                <span>Total Extra Charge</span>
                <strong>
                  Transport + GST + Extra
                </strong>
              </div>

              <div className="formula">
                <span>Extra Charge / PCS</span>
                <strong>
                  Extra Charge × Unit Rate ÷ Total
                </strong>
              </div>

              <div className="formula">
                <span>Cost Price</span>
                <strong>
                  Unit Price + Extra / PCS
                </strong>
              </div>

              <div className="formula">
                <span>Sale Price</span>
                <strong>
                  Cost × (1 + Profit%)
                </strong>
              </div>

              <div className="formula">
                <span>Default MRP</span>
                <strong>
                  Sale Price × 2
                </strong>
              </div>

            </div>

          </section>

        </div>


        {/* GENERATE SECTION */}

        <section className="generate-section">

          <div className="generate-info">

            <div className="generate-icon">
              📊
            </div>

            <div>
              <h3>
                Ready to generate your Excel file?
              </h3>

              <p>
                Your PDF will be processed and converted
                into the standardized product import format.
              </p>
            </div>

          </div>

          <button
            type="button"
            className="generate-button"
            onClick={handleGenerate}
            disabled={loading}
          >
            {loading ? (
              <>
                <span className="spinner"></span>
                Generating...
              </>
            ) : (
              <>
                Generate Excel
                <span className="button-arrow">
                  →
                </span>
              </>
            )}
          </button>

        </section>


        {/* FOOTER STATS */}

        <section className="dashboard-stats">

          <div className="stat-item">
            <span className="stat-icon">
              📄
            </span>

            <div>
              <strong>PDF</strong>
              <span>Input format</span>
            </div>
          </div>

          <div className="stat-item">
            <span className="stat-icon">
              ⚙
            </span>

            <div>
              <strong>Automated</strong>
              <span>Data processing</span>
            </div>
          </div>

          <div className="stat-item">
            <span className="stat-icon">
              📊
            </span>

            <div>
              <strong>Excel</strong>
              <span>Import ready</span>
            </div>
          </div>

          <div className="stat-item">
            <span className="stat-icon">
              ✓
            </span>

            <div>
              <strong>15 Columns</strong>
              <span>Standard format</span>
            </div>
          </div>

        </section>

      </main>


      {/* FOOTER */}

      <footer className="footer">
        <p>
          PDF2Excel Product Import Tool
        </p>

        <span>
          PDF → Product Data → Price Calculation → Excel
        </span>
      </footer>

    </div>
  );
}

export default App;