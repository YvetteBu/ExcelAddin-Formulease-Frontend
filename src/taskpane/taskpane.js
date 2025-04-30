import "../taskpane/taskpane.css";
'use strict';
const hostname = window.location.hostname;
const isDev = hostname.includes("localhost") || hostname.includes("127.0.0.1");
const API_URL = isDev 
  ? "http://localhost:3001/api/gpt" 
  : "https://excel-addin-formulease-backend.vercel.app/api/gpt";
import firebaseAuth from "../firebase.js";

if (typeof window !== "undefined") {

  // Initialize Office when the page is ready
  if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").classList.remove("hidden");

      // Authentication event listeners
      const loginButton = document.getElementById("login-button");
      const createAccountButton = document.getElementById("create-account-button");

      if (loginButton) {
        loginButton.addEventListener("click", () => {
          document.getElementById("signup-modal").style.display = "none";
          document.getElementById("login-modal").style.display = "block";
        });
      }

      if (createAccountButton) {
        createAccountButton.addEventListener("click", () => {
          document.getElementById("login-modal").style.display = "none";
          document.getElementById("signup-modal").style.display = "block";
        });
      }
      
      // Form submissions
      document.getElementById("login-form").onsubmit = handleLoginFormSubmit;
      document.getElementById("signup-form").onsubmit = handleSignupFormSubmit;

      document.getElementById("nl-generate").onclick = async () => {
        console.log("Triggered generate button");
        // Get the user's intent from the input field and trim whitespace
        let userIntent = document.getElementById("nl-input").value.trim();
        const recommendationElement = document.getElementById("recommendation");
        recommendationElement.innerHTML = "Analyzing preview of sheet and generating formula...";
        if (!userIntent) {
          recommendationElement.innerHTML = "Please enter your intention.";
          return;
        }
        // Fuzzy match userIntent tokens to headers and auto-correct
        if (typeof headers !== "undefined" && Array.isArray(headers) && headers.length > 0) {
          const loweredHeaders = headers.map(h => h.toLowerCase());
          userIntent.split(" ").forEach(word => {
            const match = loweredHeaders.find(h => h.includes(word.toLowerCase()));
            if (match) {
              userIntent = userIntent.replace(new RegExp(word, "gi"), match);
            }
          });
        }

        let usedRange, totalRows, totalCols, previewValues, headers, activeColIndex, targetHeader;
        try {
          await Excel.run(async (context) => {
            console.log("Inside Excel.run");
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            usedRange = sheet.getUsedRange();
            usedRange.load("values, address, rowCount, columnCount");
            await context.sync();
            console.log("UsedRange address:", usedRange.address);
            console.log("After sync - got usedRange");
            totalRows = usedRange.rowCount;
            totalCols = usedRange.columnCount;
            // Take first 5 data rows (excluding header), all columns up to totalCols
            const previewRows = usedRange.values.slice(1, 6); // first 5 data rows (excluding header)
            previewValues = previewRows.map(row => row.slice(0, totalCols));
            console.log("PreviewValues:", previewValues);
            const headersLocal = usedRange.values[0];
            headers = headersLocal;
            console.log("Headers:", headers);

            // Determine the active cell's column index and address for context with enhanced logging and error handling
            let excelColIndex = 0;
            try {
              const activeCell = context.workbook.getActiveCell();
              activeCell.load("columnIndex");
              activeCell.load("address");
              await context.sync();
              activeColIndex = activeCell.columnIndex;
              if (activeColIndex >= headers.length) {
                activeColIndex = 0;
              }
              targetHeader = headers[activeColIndex];
              console.log("Final targetHeader, activeColIndex:", targetHeader, activeColIndex);
              excelColIndex = activeCell.columnIndex + 1;
              console.log("Computed excelColIndex:", excelColIndex);
            } catch (err) {
              console.log("Failed to get activeCell or sync:", err);
              activeColIndex = 0;
              targetHeader = headers[0] || "Unknown";
              excelColIndex = 1;
            }
          });
        } catch (err) {
          console.log("Excel run failed: ", err);
          recommendationElement.innerHTML = "No data in sheet. Proceeding with user input only.";
          previewValues = [];
          totalRows = 0;
          totalCols = 0;
          usedRange = { address: "N/A" };
          headers = [];
          activeColIndex = 0;
          targetHeader = "";
        }

        // Diagnostic logs before building payloadData
        console.log("Preparing to build payloadData with previewValues:", previewValues);
        console.log("Headers array content:", headers);
        console.log("usedRange address:", usedRange.address);

        if (!previewValues || previewValues.length === 0) {
          recommendationElement.innerHTML = "Preview data unavailable. Please enter a more specific request.";
          return;
        }

        // Build structured payloadData object
        const payloadData = {
          instruction: userIntent,
          previewValues,
          headers,
          activeColIndex,
          targetHeader,
          usedRange: usedRange.address,
          totalRows,
          totalCols
        };
        console.log("Structured payloadData:", payloadData);
        recommendationElement.innerHTML = "Sending request to backend...";
        let response;
        // Diagnostic: check payloadData before fetch
        if (!payloadData.instruction || payloadData.instruction.length < 1) {
          console.error("Instruction is too short or undefined:", payloadData.instruction);
          recommendationElement.innerHTML = "Prompt construction failed.";
          return;
        }
        try {
          let payload;
          try {
            payload = JSON.stringify(payloadData);
          } catch (err) {
            console.error("Payload stringify failed:", err);
            recommendationElement.innerHTML = "Prompt encoding failed.";
            return;
          }
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 20000);
          response = await fetch(API_URL, {
            method: "POST",
            mode: "cors",
            credentials: "omit",
            headers: {
              "Content-Type": "application/json"
            },
            body: payload,
            signal: controller.signal
          });
          clearTimeout(timeoutId);
          console.log("Fetch complete. Status:", response.status);
          if (!response.ok) {
            const errorText = await response.text();
            console.error("Server responded with:", errorText);
            recommendationElement.innerHTML = `Server error: ${response.status}`;
            return;
          }
          let formula, explanation;
          try {
            const parsed = await response.json();
            formula = parsed.formula;
            explanation = parsed.explanation;
          } catch (err) {
            console.error("Failed to parse backend JSON:", err);
            recommendationElement.innerHTML = "Backend returned invalid response.";
            return;
          }
          // --- Begin UI logic for apply buttons and explanation from parsed JSON ---
          recommendationElement.innerHTML = "";

          // Insert error block if formula is invalid, and return early
          if (!formula || formula === "**") {
            const errorBlock = document.createElement("div");
            errorBlock.innerText = "We couldn’t generate a valid formula. Try using phrases like ‘average of troponin when result is positive’ or ‘rank by heart rate descending’.";
            errorBlock.style.color = "red";
            recommendationElement.appendChild(errorBlock);
            return;
          }

          // Patch for "sort by" prompts: convert any =SORT to =SORTBY with correct range
          let displayFormula = formula;
          if (
            payloadData.instruction &&
            /sort by/i.test(payloadData.instruction) &&
            /^=SORT\(/i.test(formula)
          ) {
            // Try to extract sorting column index and direction from the formula
            // e.g. =SORT(A2:C100, 2, -1)
            const sortMatch = formula.match(/^=SORT\(\s*([A-Z]+\d*:[A-Z]+\d*)\s*,\s*(\d+)\s*,\s*(-?1)\s*\)/i);
            if (sortMatch) {
              // Use headers and totalRows from payloadData
              const headers = payloadData.headers || [];
              const totalRows = payloadData.totalRows || 0;
              const colIndex = parseInt(sortMatch[2], 10) - 1;
              const direction = parseInt(sortMatch[3], 10);
              const colLetter = String.fromCharCode(65 + colIndex);
              const fullRange = `A2:${String.fromCharCode(65 + headers.length - 1)}${totalRows}`;
              const sortRange = `${colLetter}2:${colLetter}${totalRows}`;
              displayFormula = `=SORTBY(${fullRange}, ${sortRange}, ${direction})`;
            }
          }

          const formulaBlock = document.createElement("div");
          formulaBlock.innerText = `Formula: ${displayFormula || "**"}`;
          recommendationElement.appendChild(formulaBlock);

          const cellBlock = document.createElement("div");
          const safeTargetCell = typeof targetCell === "string" ? targetCell : "L1";
          cellBlock.innerText = `Recommended cell: ${safeTargetCell}`;
          recommendationElement.appendChild(cellBlock);

          const buttonContainer = document.createElement("div");
          buttonContainer.style.marginTop = "10px";

          const applyToSelectionBtn = document.createElement("button");
          applyToSelectionBtn.innerText = "Apply to Selected Cell";
          applyToSelectionBtn.style.marginRight = "10px";

          applyToSelectionBtn.onclick = async () => {
            const safeTargetCell = typeof targetCell === "string" ? targetCell : "L1";
            // Protect: do not insert if formula is invalid
            if (!displayFormula || displayFormula === "**") {
              console.error("Invalid formula. Skipping insert.");
              recommendationElement.innerHTML += "<div style='color:red; margin-top:8px;'>No valid formula was generated. Please revise your input.</div>";
              return;
            }
            await Excel.run(async (context) => {
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const selectedRange = context.workbook.getSelectedRange();
              selectedRange.load(["rowIndex", "columnIndex"]);
              await context.sync();

              const startRow = selectedRange.rowIndex;
              const startCol = selectedRange.columnIndex;

              const targetCellObj = sheet.getCell(startRow, startCol);
              targetCellObj.formulas = [[displayFormula.startsWith("=") ? displayFormula : "=" + displayFormula]];
              await context.sync();
            });
          };

          buttonContainer.appendChild(applyToSelectionBtn);
          recommendationElement.appendChild(buttonContainer);

          const explanationBlock = document.createElement("div");
          explanationBlock.innerText = `Explanation: ${explanation || "No explanation."}`;
          explanationBlock.style.marginTop = "10px";
          recommendationElement.appendChild(explanationBlock);
          // --- End UI logic ---

        } catch (fetchErr) {
          if (fetchErr.name === 'AbortError') {
            console.error("Fetch aborted due to timeout");
            recommendationElement.innerHTML = "Request timed out. Please try again.";
          } else {
            console.error("Fetch threw error:", fetchErr);
            recommendationElement.innerHTML = `Backend error. ${fetchErr.message || fetchErr}`;
          }
          return;
        }

      };
      // Display user info in dropdown
      const accountDropdown = document.getElementById("account-dropdown");
      if (accountDropdown) {
        const email = localStorage.getItem("userEmail") || "N/A";
        const credits = localStorage.getItem("userCredits") || "0";
        accountDropdown.innerHTML = `
          <p><strong>Email:</strong> ${email}</p>
          <p><strong>Credits:</strong> ${credits}</p>
        `;
      }

      if (info.host === Office.HostType.Excel) {
        const insertBtn = document.getElementById("insertData");
        if (insertBtn) {
          insertBtn.addEventListener("click", insertSampleData);
        }
        // Avoid runtime error if "insertData" button is missing from DOM
      }
    });
  }

  async function handleLoginFormSubmit(event) {
    event.preventDefault();
    
    const email = document.getElementById("login-email").value;
    const password = document.getElementById("login-password").value;
    const errorElement = document.getElementById("login-error");

    if (!email || !password) {
      errorElement.textContent = "Please enter both email and password";
      errorElement.style.display = "block";
      return;
    }

    try {
      const { data, error } = await firebaseAuth.signIn(email, password);
      
      if (error) {
        errorElement.textContent = error.message;
        errorElement.style.display = "block";
        return;
      }

      // Store user info in localStorage
      localStorage.setItem("userEmail", email);
      localStorage.setItem("userProfile", JSON.stringify({
        name: data.displayName || "",
        email: data.email
      }));
      
      // Update UI
      updateUserInfo(data.profile);
      closeLoginModal();
      alert("Login successful!");
    } catch (error) {
      errorElement.textContent = error.message;
      errorElement.style.display = "block";
    }
  }

  function openSignupModal() {
    document.getElementById("signup-modal").style.display = "block";
  }

  function closeSignupModal() {
    document.getElementById("signup-modal").style.display = "none";
  }

  async function handleSignupFormSubmit(event) {
    console.log("trying to sing up")

    event.preventDefault();
    
    const name = document.getElementById("signup-name").value;
    const email = document.getElementById("signup-email").value;
    const password = document.getElementById("signup-password").value;
    const confirmPassword = document.getElementById("signup-confirm-password").value;

    // Validate passwords match
    if (password !== confirmPassword) {
        alert("Passwords do not match. Please try again.");
        return;
    }

    // Validate password strength
    if (password.length < 6) {
        alert("Password must be at least 6 characters long");
        return;
    }

    try {
        // First, sign up the user
        const { data: authData, error: authError } = await firebaseAuth.signUp(email, password, { name });
        
        if (authError) {
            alert(`Signup failed: ${authError.message}`);
            return;
        }

        // Initialize user credits in Supabase
        const { error: creditError } = await firebaseAuth.initializeUserCredits(authData.user.id);
        
        if (creditError) {
            alert(`Failed to initialize credits: ${creditError.message}`);
            return;
        }

        // Store user info in localStorage
        localStorage.setItem("userEmail", email);
        localStorage.setItem("userProfile", JSON.stringify({ 
            name: name, 
            email: email,
            credits: 20
        }));
        
        // Update UI
        updateUserInfo({ name, email, credits: 20 });
        closeSignupModal();
        alert("Account created successfully! You have been given 20 free credits.");
    } catch (error) {
        alert(`Signup error: ${error.message}`);
    }
  }

  function updateUserInfo(profile) {
    const accountDropdown = document.getElementById("account-dropdown");
    if (accountDropdown) {
      accountDropdown.innerHTML = `
        <p><strong>Name:</strong> ${profile.name}</p>
        <p><strong>Email:</strong> ${profile.email}</p>
        <button onclick="handleSignOut()" style="margin-top: 10px;">Sign Out</button>
      `;
    }
  }

  // Make handleSignOut available globally
  window.handleSignOut = async function() {
    try {
      const { error } = await firebaseAuth.signOut();
      if (error) throw error;
      
      // Clear localStorage
      localStorage.removeItem("userEmail");
      localStorage.removeItem("userProfile");
      
      // Update UI
      const accountDropdown = document.getElementById("account-dropdown");
      if (accountDropdown) {
        accountDropdown.innerHTML = `
          <p><strong>Status:</strong> Not logged in</p>
        `;
      }
      
      alert("Signed out successfully");
    } catch (error) {
      alert(`Sign out error: ${error.message}`);
    }
  };

  async function insertSampleData() {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = [
          ["Product", "Quantity", "Price"],
          ["Widget A", 10, 19.99],
          ["Widget B", 15, 29.99],
          ["Widget C", 20, 39.99]
        ];
        range.format.autofitColumns();
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }
}