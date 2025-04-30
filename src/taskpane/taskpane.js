import "../taskpane/taskpane.css";
'use strict';
const isDev = window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1') || window.location.hostname.includes('excel.officeapps.live.com');
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
        const userIntent = document.getElementById("nl-input").value.trim();
        const recommendationElement = document.getElementById("recommendation");
        recommendationElement.innerHTML = "Analyzing preview of sheet and generating formula...";
        if (!userIntent) {
          recommendationElement.innerHTML = "Please enter your intention.";
          return;
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
            previewValues = usedRange.values.slice(1, 6); // first 5 rows, all columns
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

        // Diagnostic logs before building userPrompt
        console.log("Preparing to build userPrompt with previewValues:", previewValues);
        console.log("Headers array content:", headers);
        console.log("usedRange address:", usedRange.address);

        if (!previewValues || previewValues.length === 0) {
          recommendationElement.innerHTML = "Preview data unavailable. Please enter a more specific request.";
          return;
        }

        const userPrompt = `
You are an expert Excel formula assistant.

User's request: "${userIntent}" (use this as the main instruction)

- Preview (first 10 rows, range ${usedRange.address}, size: ${totalRows}x${totalCols}): ${JSON.stringify(previewValues)}
- The active column is: ${targetHeader} (column index ${excelColIndex})
- Use the active column (index ${excelColIndex}) as the default target for calculations unless the user specifies otherwise.

Instructions:
- Row 1 contains headers.
- Output must be a single Excel formula.
- Also suggest the TargetCell.
- Format your answer like this (no extra text):
  Formula: =...
  TargetCell: ...
  Explanation: ...
- DO NOT use the SORT function. Use SORTBY instead.
- DO NOT reference headers in formulas.
- Ensure compatibility with common Excel versions.
`;
        console.log("Constructed userPrompt:", userPrompt);

        let response;
        // Diagnostic: check prompt before fetch
        if (!userPrompt || userPrompt.length < 10) {
          console.error("Prompt is too short or undefined:", userPrompt);
          recommendationElement.innerHTML = "Prompt construction failed.";
          return;
        }
        try {
          let payload;
          try {
            payload = JSON.stringify({ prompt: userPrompt });
          } catch (err) {
            console.error("Prompt stringify failed:", err);
            recommendationElement.innerHTML = "Prompt encoding failed.";
            return;
          }
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 10000);
          response = await fetch(API_URL, {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: payload,
            signal: controller.signal
          });
          clearTimeout(timeoutId);
          console.log("Fetch complete. Status:", response.status);
          if (!response.ok) {
            recommendationElement.innerHTML = `Server error: ${response.status}`;
            return;
          }
          const res = await response.text();
          console.log("Received response:", res);
          if (res.includes("Method Not Allowed")) {
            recommendationElement.innerHTML = "Backend rejected the request method. Please contact developer.";
            return;
          }
          if (!res.trim()) {
            recommendationElement.innerHTML = "Empty response from backend. Please try again.";
            return;
          }
          // recommendationElement.innerHTML = res;

          // --- Begin UI logic for apply buttons and explanation from plain res ---
          const reply = res.trim();

          const formulaMatch = reply.match(/Formula:\s*=(.+?)(?:\r?\n|$)/s);
          const targetCellMatch = reply.match(/TargetCell:\s*([A-Z]+\d+)/);
          const explanationMatch = reply.match(/Explanation:\s*([\s\S]*?)(?:\r?\n[A-Z][a-z]+:|$)/);

          // Patch logical AND syntax fix for FILTER usage
          if (formulaMatch && formulaMatch[1] && formulaMatch[1].includes("FILTER(")) {
            let rawFormula = formulaMatch[1];
            rawFormula = rawFormula.replace(/FILTER\(([^,]+),\s*\(([^)]+)\)\s*\*\s*\(([^)]+)\)\)/, 
              (match, range, cond1, cond2) => `FILTER(${range}, (${cond1})*(${cond2}))`);
            formulaMatch[1] = rawFormula;
          }

          const formula = formulaMatch ? formulaMatch[1] : null;
          const isRangeFormula = formula && /^(SORTBY|FILTER)\(/i.test(formula.trim());
          const targetCell = targetCellMatch ? targetCellMatch[1] : "A1";
          const explanation = explanationMatch ? explanationMatch[1].trim() : "No explanation.";

          if (!formula) {
            recommendationElement.innerHTML = "Failed to extract formula. Please try rephrasing your request.";
            return;
          }

          if (formula) {
            const formulaBlock = document.createElement("div");
            formulaBlock.innerText = `Formula: ${formula}`;
            recommendationElement.appendChild(formulaBlock);

            const cellBlock = document.createElement("div");
            cellBlock.innerText = `Recommended cell: ${targetCell}`;
            recommendationElement.appendChild(cellBlock);

            const buttonContainer = document.createElement("div");
            buttonContainer.style.marginTop = "10px";

            const applyToSelectionBtn = document.createElement("button");
            applyToSelectionBtn.innerText = "Apply to Selected Cell";
            applyToSelectionBtn.style.marginRight = "10px";


            // Formula application logic: minimal version, Excel handles spill/display natively
            applyToSelectionBtn.onclick = async () => {
              await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const selectedRange = context.workbook.getSelectedRange();
                selectedRange.load(["rowIndex", "columnIndex"]);
                await context.sync();

                const startRow = selectedRange.rowIndex;
                const startCol = selectedRange.columnIndex;
                // Write formula to a safe spill zone outside used range to avoid #CALC! errors
                const usedRange = sheet.getUsedRange();
                usedRange.load("rowCount, columnCount");
                await context.sync();
                const targetCell = sheet.getCell(0, usedRange.columnCount + 2);
                const spillClearRange = sheet.getRangeByIndexes(0, usedRange.columnCount + 2, usedRange.rowCount + 10, usedRange.columnCount);
                spillClearRange.clear(); // Ensure all spill cells are truly empty
                await context.sync();

                targetCell.formulas = [[formula.startsWith("=") ? formula : "=" + formula]];
                await context.sync();
              });
            };


            buttonContainer.appendChild(applyToSelectionBtn);


            recommendationElement.appendChild(buttonContainer);
          }

          const explanationBlock = document.createElement("div");
          explanationBlock.innerText = `Explanation: ${explanation}`;
          explanationBlock.style.marginTop = "10px";
          recommendationElement.appendChild(explanationBlock);
          // --- End UI logic for apply buttons and explanation from plain res ---
          
        } catch (fetchErr) {
          console.error("Fetch threw error:", fetchErr);
          recommendationElement.innerHTML = `Backend error. ${fetchErr.message || fetchErr}`;
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