import "../taskpane/taskpane.css";
'use strict';
const API_URL = window.location.hostname.includes('localhost') 
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
        // Get the user's intent from the input field and trim whitespace
        const userIntent = document.getElementById("nl-input").value.trim();
        const recommendationElement = document.getElementById("recommendation");
        recommendationElement.innerHTML = "Analyzing preview of sheet and generating formula...";
        if (!userIntent) {
          recommendationElement.innerHTML = "Please enter your intention.";
          return;
        }

        let usedRange, totalRows, totalCols, previewValues;

        try {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            usedRange = sheet.getUsedRange();
            usedRange.load("values, address, rowCount, columnCount");
            await context.sync();
            totalRows = usedRange.rowCount;
            totalCols = usedRange.columnCount;
            previewValues = usedRange.values.slice(1, 11); // preview first 10 rows
          });
        } catch (err) {
          console.log("Excel run failed: ", err);
          recommendationElement.innerHTML = "No data in sheet. Proceeding with user input only.";
          previewValues = [];
          totalRows = 0;
          totalCols = 0;
          usedRange = { address: "N/A" };
        }

        const userPrompt = `
        You are an expert Excel formula assistant.

        User's request: "${userIntent}" (use this as the main instruction)

        - Preview (first 10 rows, range ${usedRange.address}, size: ${totalRows}x${totalCols}):${JSON.stringify(previewValues)}

        Instructions:
        - Row 1 contains headers.
        - Output must be a single Excel formula.
        - Also suggest the TargetCell.
        - Format your answer like this (no extra text):
          Formula: =...
          TargetCell: ...
          Explanation: ...
        - Be careful to avoid formulas that throw errors in Excel.
        - DO NOT use the SORT function. Use SORTBY instead.
        - DO NOT reference headers in formulas.
        - Ensure the formula is compatible with common Excel versions.
        `;

        let response;
        try {
          console.log("Requesting backend from:", API_URL);
          if (!API_URL.includes("localhost")) console.warn("Using production backend, ensure taskpane.js and taskpane.css are correctly deployed.");
          response = await fetch(API_URL, {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            mode: "cors",  // explicitly allow CORS
            body: JSON.stringify({ prompt: userPrompt })
          });

          console.log("Response status:", response.status);
          if (!response.ok) {
            recommendationElement.innerHTML = `Server error: ${response.status}`;
            return;
          }

          const res = await response.text();
          console.log("Received response:", res);
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
          const targetCell = targetCellMatch ? targetCellMatch[1] : "A1";
          const explanation = explanationMatch ? explanationMatch[1].trim() : "No explanation.";

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


            applyToSelectionBtn.onclick = async () => {
              await Excel.run(async (context) => {
                try {
                  const sheet = context.workbook.worksheets.getActiveWorksheet();
                  const usedRange = sheet.getUsedRange();
                  usedRange.load("rowCount, columnCount");
                  await context.sync();

                  const selectedRange = context.workbook.getSelectedRange();
                  selectedRange.load(["rowIndex", "columnIndex"]);
                  await context.sync();

                  const startRow = selectedRange.rowIndex;
                  const startCol = selectedRange.columnIndex;
                  const totalRows = usedRange.rowCount - startRow;

                  // New logic: only allow aggregate functions to insert into a single cell and insert the computed result
                  const aggregateFunctions = ["AVERAGE(", "SUM(", "MAX(", "MIN(", "COUNT("];
                  const normalizedFormula = formula.replace(/\s+/g, "").toUpperCase();
                  const isAggregate = aggregateFunctions.some(fn => normalizedFormula.includes(fn));
                  // Allow aggregate formulas even if they are wrapped (e.g., AVERAGE(FILTER(...)))
                  const isSingleCellFormula = isAggregate;

                  if (isSingleCellFormula) {
                    // For aggregate functions, insert the formula, wait for calculation, then overwrite with value
                    const resultRange = sheet.getRangeByIndexes(startRow, startCol, 1, 1);
                    resultRange.formulas = [[formula]];
                    await context.sync();

                    // Load the calculated value and overwrite if successful, else fallback to formula
                    resultRange.load("values");
                    await context.sync();
                    const finalValue = resultRange.values[0][0];
                    if (finalValue !== undefined && !(typeof finalValue === "string" && finalValue.startsWith("#"))) {
                      resultRange.values = [[finalValue]];
                      await context.sync();
                      console.log("Final value inserted into cell:", finalValue);
                    } else {
                      resultRange.formulas = [["=" + formula]];
                      await context.sync();
                      console.warn("Formula fallback inserted due to unresolved value.");
                    }
                  } else {
                    // fallback to inserting raw formula into the selected range, but try to evaluate and replace with value if possible
                    const fallbackRange = sheet.getRangeByIndexes(startRow, startCol, 1, 1);
                    fallbackRange.formulas = [["=" + formula]];
                    await context.sync();
                    fallbackRange.load("values");
                    await context.sync();
                    const computedValue = fallbackRange.values[0][0];
                    if (computedValue !== undefined && !(typeof computedValue === "string" && computedValue.startsWith("#"))) {
                      fallbackRange.values = [[computedValue]];
                      await context.sync();
                    }
                  }
                } catch (error) {
                  console.error("Failed to apply formula:", error);
                }
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
          
        } catch (err) {
          console.log("Fetch failed: ", err);
          recommendationElement.innerHTML = `Request error: ${err.message || err}`;
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