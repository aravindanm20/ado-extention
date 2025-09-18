VSS.init({
    explicitNotifyLoaded: true,
    usePlatformStyles: true
});

VSS.ready(function () {
    // Keep a reference to the AI data to use during save
    let generatedAiData = null;

    // UI elements
    const generateBtn = document.getElementById('generateBtn');
    const saveBtn = document.getElementById('saveBtn');
    const useCaseInput = document.getElementById('useCaseInput');
    const epicTitle = document.getElementById('epicTitle');
    const epicDescription = document.getElementById('epicDescription');
    const epicAcceptanceCriteria = document.getElementById('epicAcceptanceCriteria');

    // 1. Event listener for the "Generate" button
    generateBtn.addEventListener('click', async () => {
        const userDescription = useCaseInput.value;
        if (!userDescription) {
            alert('Please provide a use case description.');
            return;
        }
        payload={
            mode: "ADOPLUGIN_TEST",
            useCaseIdentifier: "ADOPLUGIN_TEST@ASCENDION@AAVA_DOMAIN@AAVA_PROJECT@AAVA_TEAM",
            userSignature: "aravindan.moorthy@ascendion.com",
            prompt: userDescription,
            promptOverride: false
        }

        const aiApiUrl = 'https://avaplus-internal.avateam.io/v1/api/instructions/ava/force/individualAgent/execute';

        try {
            generateBtn.disabled = true;
            generateBtn.textContent = 'Generating...';

            const response = await fetch(aiApiUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json',
                    'access-key' : 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJ2b1B5VG5PNFJPS1dYYkRiOWY2YjdSZlZaY0t5cDJkbkNiZjlsX1V1azJJIiwiaWF0IjoxNzU4MjE3ODUyLCJleHAiOjE3NTgyMjE3MDksImF1ZCI6IjAwMDAwMDAzLTAwMDAtMDAwMC1jMDAwLTAwMDAwMDAwMDAwMCIsInRpZCI6ImQ3NzU4ZThmLTFkZjMtNDg5Zi04NmI1LWEyMjU0ZjU1ZjljYyIsImFwcGlkIjoiMjk2Zjk0ZDgtMWMzNy00ZjNlLTlkZjctMTg3ZmY3ZWZmZmFkIiwidW5pcXVlX25hbWUiOiJhcmF2aW5kYW4ubW9vcnRoeUBhc2NlbmRpb24uY29tIn0.U7tvqC9QeoyxdgNHkjWjHfPHgzrkEP63Wk-CF-IvLPYAtyc-ZNx5jxWHasIHOwesRMU2C2_NVdsrPjWUWuBH5JUq7u1x9Dv17iYd9nouUS4nRFwmW2pjPpu4x23db7XCYXliZPVDMZ6wDMW4uWF3WIVxkjrB6qWwnJXXfOttDQ2A05CeOkJ8LENCxJagWhfyUbCxfTGvjTr7kJihiAblj5bsgM7J6mhNxqdSnzfZ8ePCXyfAyHU6Wt2uQiDLBqO6U3IMN1HxG7Bue8sm1sH1kGSvRmbq7jzMocetD-QMmVQSKLz-ZBuTi2LcVgB0ES8juuiIFDJAfS1nqDliC8WvAg'
                 },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                throw new Error(`AI API Error: ${response.statusText}`);
            }

            const fullApiResponse = await response.json();

            // **STEP 1: Extract the text string from the nested response**
            const rawText = fullApiResponse.response.choices[0].text;

            // **STEP 2: Clean the string by removing markdown backticks and 'json' identifier**
            const cleanedText = rawText.replace(/```json\s*|```/g, '').trim();

            // **STEP 3: Parse the cleaned string into a JavaScript object**
            const aiData = JSON.parse(cleanedText);
            generatedAiData = aiData; // Store for the save function

            // **STEP 4: Populate the form with the rich data from the AI**
            epicTitle.value = aiData.Title || '';

            // Combine Description, Dependencies, and Notes into one field for ADO
            let fullDescriptionHtml = aiData.Description ? aiData.Description.replace(/\\n/g, '<br>') : '';
            if (aiData.Dependencies && aiData.Dependencies.length > 0) {
                fullDescriptionHtml += `<br><br><b>Dependencies:</b><ul>${aiData.Dependencies.map(d => `<li>${d}</li>`).join('')}</ul>`;
            }
            if (aiData.Notes) {
                fullDescriptionHtml += `<br><br><b>Notes:</b><br>${aiData.Notes.replace(/\\n/g, '<br>')}`;
            }
            epicDescription.value = fullDescriptionHtml;

            // Format Acceptance Criteria as an HTML list
            if (aiData.AcceptanceCriteria && Array.isArray(aiData.AcceptanceCriteria)) {
                epicAcceptanceCriteria.value = `<ul>${aiData.AcceptanceCriteria.map(ac => `<li>${ac}</li>`).join('')}</ul>`;
            }

            saveBtn.disabled = false;

        } catch (error) {
            console.error('Failed to call or parse AI endpoint response:', error);
            alert(`Error generating details: ${error.message}`);
        } finally {
            generateBtn.disabled = false;
            generateBtn.textContent = 'Generate Details';
        }
    });

    // 2. Event listener for the "Save" button
    saveBtn.addEventListener('click', () => {
        if (!generatedAiData) {
            alert("No AI data has been generated yet.");
            return;
        }
        
        const context = VSS.getWebContext();
        const projectName = context.project.name;

        // Helper function to map string priority to ADO's integer value
        const mapPriority = (priorityString) => {
            const priorityMap = { "high": 1, "medium": 2, "low": 3 };
            return priorityMap[priorityString.toLowerCase()] || 2; // Default to Medium
        };

        const patchDocument = [
            { "op": "add", "path": "/fields/System.Title", "value": epicTitle.value },
            { "op": "add", "path": "/fields/System.Description", "value": epicDescription.value },
            { "op": "add", "path": "/fields/Microsoft.VSTS.Common.AcceptanceCriteria", "value": epicAcceptanceCriteria.value },
            { "op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": mapPriority(generatedAiData.Priority) },
            { "op": "add", "path": "/fields/System.AreaPath", "value": context.project.name },
            { "op": "add", "path": "/fields/System.IterationPath", "value": context.project.name }
        ];

        VSS.getService(VSS.ServiceIds.WorkItemTracking).then(function (witApi) {
            saveBtn.disabled = true;
            saveBtn.textContent = 'Saving...';
            
            witApi.createWorkItem(patchDocument, projectName, "Epic").then(
                function(workItem) {
                    alert(`Successfully created Epic #${workItem.id}`);
                    VSS.getConfiguration().dialog.close();
                },
                function(error) {
                    console.error(error);
                    alert(`Failed to create Epic: ${error.message}`);
                    saveBtn.disabled = false;
                    saveBtn.textContent = 'Save Epic';
                }
            );
        });
    });

    VSS.notifyLoadSucceeded();
});