// modules/word.js - ENHANCED WORD MODULE WITH BETTER DETECTION
const { exec } = require('child_process');
const path = require('path');
const os = require('os');

const DESKTOP_PATH = path.join(os.homedir(), 'Desktop');

// Enhanced Word handling with Claude API support
async function handleSmartWord(message, claudeModule) {
    console.log('📝 Enhanced Word processing with better detection...');
    
    try {
        // Step 1: Comprehensive Word detection
        const wordStatus = await detectWordComprehensively();
        console.log('🔍 Word detection result:', wordStatus);
        
        // Step 2: Try different approaches based on what's available
        if (wordStatus.comAvailable) {
            console.log('✅ Word COM available, attempting automation...');
            const result = await attemptWordAutomationEnhanced(message, claudeModule);
            return {
                message: '📝 Word automation attempted',
                details: result
            };
        } else if (wordStatus.executableFound) {
            console.log('⚠️ Word executable found but COM blocked, using file-based approach...');
            const result = await createDocumentAndOpenWord(message, claudeModule);
            return {
                message: '📝 Word document created and opened',
                details: result
            };
        } else {
            console.log('❌ Word not available, creating templates...');
            const result = await createAdvancedDocumentTemplate(message, claudeModule);
            return {
                message: '📝 Document templates created (Word not available)',
                details: result
            };
        }
        
    } catch (error) {
        console.log('❌ Word processing error, falling back to templates...');
        const result = await createAdvancedDocumentTemplate(message, claudeModule);
        return {
            message: '📝 Word automation failed - templates created instead',
            details: `⚠️ Word error: ${error.message}\n\n${result}`
        };
    }
}

// Comprehensive Word detection
function detectWordComprehensively() {
    return new Promise((resolve) => {
        console.log('🔍 Running comprehensive Word detection...');
        
        const results = {
            executableFound: false,
            executablePath: null,
            registryFound: false,
            registryPath: null,
            comAvailable: false,
            processRunning: false,
            version: null
        };
        
        let testsCompleted = 0;
        const totalTests = 4;
        
        // Test 1: Check if Word executable exists in PATH
        exec('where winword', (error, stdout, stderr) => {
            if (!error && stdout.trim()) {
                results.executableFound = true;
                results.executablePath = stdout.trim().split('\n')[0];
                console.log('✅ Word executable found in PATH:', results.executablePath);
            }
            testsCompleted++;
            if (testsCompleted === totalTests) finalizeResults();
        });
        
        // Test 2: Check registry for Word installation
        exec('reg query "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE" /ve 2>nul', (error, stdout, stderr) => {
            if (!error && stdout.includes('REG_SZ')) {
                const match = stdout.match(/REG_SZ\s+(.+)$/m);
                if (match) {
                    results.registryFound = true;
                    results.registryPath = match[1].trim();
                    console.log('✅ Word found in registry:', results.registryPath);
                }
            }
            testsCompleted++;
            if (testsCompleted === totalTests) finalizeResults();
        });
        
        // Test 3: Check if Word is currently running
        exec('tasklist /FI "IMAGENAME eq WINWORD.EXE" /FO CSV | find "winword.exe"', (error, stdout, stderr) => {
            if (!error && stdout.includes('winword.exe')) {
                results.processRunning = true;
                console.log('🔄 Word is currently running');
            }
            testsCompleted++;
            if (testsCompleted === totalTests) finalizeResults();
        });
        
        // Test 4: Test COM object creation (most important for automation)
        const comTestScript = `
try {
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
    $word.Visible = $false
    $version = $word.Version
    $word.Quit()
    Write-Host "COM_SUCCESS:$version"
} catch {
    Write-Host "COM_FAILED:$($_.Exception.Message)"
}`;
        
        const tempScript = path.join(os.tmpdir(), `word_com_test_${Date.now()}.ps1`);
        require('fs').writeFileSync(tempScript, comTestScript);
        
        exec(`powershell -ExecutionPolicy Bypass -File "${tempScript}"`, (error, stdout, stderr) => {
            try { require('fs').unlinkSync(tempScript); } catch (e) {}
            
            if (stdout && stdout.includes('COM_SUCCESS')) {
                results.comAvailable = true;
                const versionMatch = stdout.match(/COM_SUCCESS:(.+)/);
                if (versionMatch) {
                    results.version = versionMatch[1].trim();
                    console.log('✅ Word COM object working, version:', results.version);
                }
            } else {
                console.log('❌ Word COM object failed:', stdout || stderr || error?.message);
            }
            
            testsCompleted++;
            if (testsCompleted === totalTests) finalizeResults();
        });
        
        function finalizeResults() {
            // Determine best path for Word
            if (!results.executableFound && results.registryFound) {
                results.executableFound = true;
                results.executablePath = results.registryPath;
            }
            
            console.log('📊 Word detection summary:', results);
            resolve(results);
        }
        
        // Timeout after 10 seconds
        setTimeout(() => {
            if (testsCompleted < totalTests) {
                console.log('⏰ Word detection timeout, using partial results');
                finalizeResults();
            }
        }, 10000);
    });
}

// Enhanced Word automation with better error handling
async function attemptWordAutomationEnhanced(message, claudeModule) {
    // First, generate smart content if Claude is available
    let docDetails = parseDocumentRequest(message);
    
    try {
        if (claudeModule) {
            console.log('🤖 Generating smart document content...');
            const contentResult = await claudeModule.generateDocumentContent(message, docDetails.type);
            
            if (contentResult.success) {
                docDetails.content = contentResult.content;
                docDetails.aiGenerated = true;
                docDetails.generatedBy = contentResult.generatedBy;
                console.log('✅ Smart content generated by Claude API');
            } else {
                console.log('⚠️ Claude content generation failed, using template');
            }
        }
    } catch (error) {
        console.log('⚠️ Smart content generation error:', error.message);
    }
    
    return new Promise((resolve) => {
        const scriptContent = `
# Enhanced Word Automation Script
Write-Host "📝 Starting enhanced Word automation..."

try {
    # Create Word Application object with detailed error handling
    Write-Host "🔧 Creating Word application object..."
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
    
    # Configure Word settings
    $word.Visible = $true
    $word.DisplayAlerts = $false
    Write-Host "✅ Word application created successfully (Version: $($word.Version))"
    
    # Create new document
    Write-Host "📄 Creating new document..."
    $doc = $word.Documents.Add()
    Write-Host "✅ New document created"
    
    # Get selection object for content insertion
    $selection = $word.Selection
    
    # Insert title if provided
    if ("${docDetails.title}" -ne "") {
        Write-Host "📝 Adding document title..."
        $selection.Font.Name = "Arial"
        $selection.Font.Size = 18
        $selection.Font.Bold = $true
        $selection.ParagraphFormat.Alignment = 1  # Center alignment
        $selection.TypeText("${docDetails.title}")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        
        # Reset formatting for body
        $selection.Font.Size = 12
        $selection.Font.Bold = $false
        $selection.ParagraphFormat.Alignment = 0  # Left alignment
        Write-Host "✅ Title added and formatted"
    }
    
    # Insert main content with proper formatting
    Write-Host "📄 Adding document content..."
    $contentLines = @"
${docDetails.content.replace(/"/g, '""')}
"@ -split "\\n"
    
    foreach ($line in $contentLines) {
        if ($line.Trim() -ne "") {
            # Check for formatting hints
            if ($line -like "*[HEADING]*") {
                $cleanLine = $line -replace "\\[HEADING\\]", ""
                $selection.Font.Bold = $true
                $selection.Font.Size = 14
                $selection.TypeText($cleanLine.Trim())
                $selection.Font.Bold = $false
                $selection.Font.Size = 12
            } else {
                $selection.TypeText($line.Trim())
            }
            $selection.TypeParagraph()
        } else {
            $selection.TypeParagraph()  # Empty line
        }
    }
    
    ${docDetails.aiGenerated ? `
    # Add AI generation note
    $selection.TypeParagraph()
    $selection.Font.Size = 10
    $selection.Font.Italic = $true
    $selection.TypeText("---")
    $selection.TypeParagraph()
    $selection.TypeText("Generated by ${docDetails.generatedBy} on $(Get-Date -Format 'MMM dd, yyyy HH:mm')")
    $selection.Font.Italic = $false
    $selection.Font.Size = 12
    ` : ''}
    
    # Move cursor to beginning of document
    $selection.HomeKey(6) | Out-Null  # wdStory = 6
    
    Write-Host "✅ Content added successfully"
    Write-Host "📊 Document type: ${docDetails.type}"
    Write-Host "🎉 Word document ready for editing!"
    
    # Optional: Save document to desktop
    try {
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $fileName = "${docDetails.type.replace(/\s/g, '_')}_$(Get-Date -Format 'yyyyMMdd_HHmmss').docx"
        $fullPath = Join-Path $desktopPath $fileName
        $doc.SaveAs2($fullPath)
        Write-Host "💾 Document saved to Desktop as: $fileName"
    } catch {
        Write-Host "⚠️ Could not auto-save document: $($_.Exception.Message)"
    }
    
    } catch {
        Write-Error "AUTOMATION_FAILED: $($_.Exception.Message)"
        }
        // End of PowerShell script content
        `;
            // Execute the PowerShell script here (implementation needed)
            // For now, resolve with scriptContent for debugging
            resolve({ script: scriptContent, docDetails });
        });
    }