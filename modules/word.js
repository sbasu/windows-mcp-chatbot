// modules/word.js - COMPLETE ENHANCED WORD MODULE WITH BETTER DETECTION
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
    $contentText = @"
${docDetails.content.replace(/"/g, '""').replace(/\n/g, '" + "`r`n" + @"')}
"@
    
    # Split content into lines and process formatting
    $contentLines = $contentText -split "\\r?\\n"
    
    foreach ($line in $contentLines) {
        $trimmedLine = $line.Trim()
        if ($trimmedLine -ne "") {
            # Check for formatting hints
            if ($trimmedLine -like "*[HEADING]*") {
                $cleanLine = $trimmedLine -replace "\\[HEADING\\]", "" -replace "\\[/HEADING\\]", ""
                $selection.Font.Bold = $true
                $selection.Font.Size = 14
                $selection.TypeText($cleanLine.Trim())
                $selection.Font.Bold = $false
                $selection.Font.Size = 12
            } else {
                $selection.TypeText($trimmedLine)
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
    
    Write-Host "🎉 WORD_AUTOMATION_SUCCESS: Document ready for editing!"
    
} catch {
    Write-Error "AUTOMATION_FAILED: $($_.Exception.Message)"
    
    # Detailed error analysis
    Write-Host "🔍 Error Analysis:"
    if ($_.Exception.Message -like "*0x800A01A8*") {
        Write-Host "   • Word COM object blocked by security policy (Zscaler/Group Policy)"
    } elseif ($_.Exception.Message -like "*0x80080005*") {
        Write-Host "   • Server execution failed - Word may not be properly installed"
    } elseif ($_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access*") {
        Write-Host "   • Permission denied - try running as administrator"
    } elseif ($_.Exception.Message -like "*RPC*") {
        Write-Host "   • RPC server unavailable - Word service may be disabled"
    } else {
        Write-Host "   • $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "💡 Troubleshooting suggestions:"
    Write-Host "   1. Close any existing Word instances and try again"
    Write-Host "   2. Try opening Word manually from Start menu first"
    Write-Host "   3. Run this script as administrator"
    Write-Host "   4. Check if corporate security (Zscaler) is blocking COM objects"
    Write-Host "   5. Verify Word is properly licensed and activated"
}`;

        const tempScript = path.join(os.tmpdir(), `word_automation_enhanced_${Date.now()}.ps1`);
        require('fs').writeFileSync(tempScript, scriptContent);

        console.log('🚀 Executing enhanced Word automation script...');

        exec(`powershell -ExecutionPolicy Bypass -File "${tempScript}"`, (error, stdout, stderr) => {
            // Clean up temp script
            try {
                require('fs').unlinkSync(tempScript);
            } catch (e) {
                console.log('Could not delete temp script:', e.message);
            }

            const output = stdout || '';
            const errorOutput = stderr || '';

            if (output.includes('WORD_AUTOMATION_SUCCESS')) {
                resolve(`✅ WORD AUTOMATION SUCCESS!

${output}

📄 Document Details:
• Type: ${docDetails.type}
• Title: ${docDetails.title}
• AI Generated: ${docDetails.aiGenerated ? 'Yes' : 'No'}
${docDetails.aiGenerated ? `• Generated By: ${docDetails.generatedBy}` : ''}
• Status: Ready for editing in Word

🎉 SUCCESS: Word should now be open with your document ready to edit!

💡 Next Steps:
1. Review the document content
2. Make any needed adjustments
3. Save with your preferred filename
4. The document was auto-saved to Desktop as backup`);

            } else if (output.includes('AUTOMATION_FAILED') || error) {
                resolve(`❌ WORD AUTOMATION FAILED

Detailed Output:
${output}

Error Details:
${errorOutput}

🛡️ This failure is likely due to:
• Corporate security restrictions (Zscaler blocking COM objects)
• Word not properly installed or configured
• Insufficient permissions or licensing issues

📝 ALTERNATIVE SOLUTIONS:
Don't worry! The system will now create advanced document templates as backup.

💡 MANUAL STEPS TO TRY:
1. Open Word manually from Start menu
2. Look for any auto-saved document on Desktop
3. Or use the template files that will be created as fallback
4. Try running the application as administrator`);
            } else {
                resolve(`⚠️ WORD AUTOMATION RESULT UNCLEAR

Output: ${output}
Error: ${errorOutput}

The automation script completed but results are uncertain.

🔍 Please check:
1. If Word opened with your document
2. Desktop for any saved documents (.docx files)
3. If Word is running but empty, the automation may have partially worked

If Word didn't open properly, template files will be created as backup.`);
            }
        });
    });
}

// Create document and attempt to open Word (fallback method)
async function createDocumentAndOpenWord(message, claudeModule) {
    try {
        // Generate content first
        let docDetails = parseDocumentRequest(message);
        
        if (claudeModule) {
            try {
                console.log('🤖 Generating smart content for file-based approach...');
                const contentResult = await claudeModule.generateDocumentContent(message, docDetails.type);
                if (contentResult.success) {
                    docDetails.content = contentResult.content;
                    docDetails.aiGenerated = true;
                    docDetails.generatedBy = contentResult.generatedBy;
                    console.log('✅ Smart content generated for document file');
                }
            } catch (error) {
                console.log('Claude content generation failed:', error.message);
            }
        }
        
        // Create Word-compatible document
        const timestamp = Date.now();
        const fileName = `${docDetails.type.replace(/\s/g, '_')}_${timestamp}.rtf`;
        const filePath = path.join(DESKTOP_PATH, fileName);
        
        // Create RTF content that Word can open
        const rtfContent = createAdvancedRTF(docDetails);
        require('fs').writeFileSync(filePath, rtfContent);
        console.log('📄 RTF document created:', fileName);
        
        // Try to open the document in Word
        return new Promise((resolve) => {
            const command = `start "" "${filePath}"`;
            
            exec(command, (error, stdout, stderr) => {
                if (error) {
                    resolve(`📝 DOCUMENT CREATED BUT WORD OPENING FAILED:

✅ Document created successfully: ${fileName}
📁 Location: Desktop
❌ Auto-opening failed: ${error.message}

💡 MANUAL STEPS:
1. Go to Desktop
2. Double-click: ${fileName}
3. It should open in Microsoft Word for editing

📄 DOCUMENT DETAILS:
• Type: ${docDetails.type}
• Title: ${docDetails.title}
• AI Generated: ${docDetails.aiGenerated ? 'Yes' : 'No'}
${docDetails.aiGenerated ? `• Generated By: ${docDetails.generatedBy}` : ''}
• Format: Rich Text Format (.rtf) - Compatible with Word

🔧 ALTERNATIVE OPENING METHODS:
• Right-click file → "Open with" → Microsoft Word
• Open Word first, then File → Open → select the document
• Copy content from file and paste into new Word document

The document is ready - just needs manual opening!`);
                } else {
                    resolve(`✅ DOCUMENT CREATED AND OPENED SUCCESSFULLY:

📄 Document: ${fileName}
📁 Location: Desktop  
🚀 Status: Should be opening in Word now

📊 DOCUMENT DETAILS:
• Type: ${docDetails.type}
• Title: ${docDetails.title}
• AI Generated: ${docDetails.aiGenerated ? 'Yes' : 'No'}
${docDetails.aiGenerated ? `• Generated By: ${docDetails.generatedBy}` : ''}
• Format: Rich Text Format (.rtf) - Fully Word compatible

🎉 Success! The document should now be open in Microsoft Word for editing.

💡 What to expect:
• Word should launch automatically
• Document will open with proper formatting
• Content will be ready for editing and customization`);
                }
            });
        });
        
    } catch (error) {
        throw new Error(`Document creation failed: ${error.message}`);
    }
}

// Create advanced document template with multiple formats
async function createAdvancedDocumentTemplate(message, claudeModule) {
    let docDetails = parseDocumentRequest(message);
    
    // Try to get Claude-generated content
    if (claudeModule) {
        try {
            console.log('🤖 Generating smart content for templates...');
            const contentResult = await claudeModule.generateDocumentContent(message, docDetails.type);
            if (contentResult.success) {
                docDetails.content = contentResult.content;
                docDetails.aiGenerated = true;
                docDetails.generatedBy = contentResult.generatedBy;
                console.log('✅ Smart content generated for templates');
            }
        } catch (error) {
            console.log('Claude content generation failed, using template:', error.message);
        }
    }
    
    const timestamp = Date.now();
    
    try {
        const results = [];
        
        // 1. Create RTF file (best for Word)
        console.log('📄 Creating RTF template...');
        const rtfContent = createAdvancedRTF(docDetails);
        const rtfPath = path.join(DESKTOP_PATH, `${docDetails.type}_${timestamp}.rtf`);
        require('fs').writeFileSync(rtfPath, rtfContent);
        results.push(`📄 ${path.basename(rtfPath)} - Rich Text Format (opens in Word)`);
        
        // 2. Create DOCX-compatible HTML
        console.log('🌐 Creating HTML template...');
        const htmlContent = createWordCompatibleHTML(docDetails);
        const htmlPath = path.join(DESKTOP_PATH, `${docDetails.type}_${timestamp}.html`);
        require('fs').writeFileSync(htmlPath, htmlContent);
        results.push(`🌐 ${path.basename(htmlPath)} - HTML (can be opened in Word via File → Open)`);
        
        // 3. Create plain text template
        console.log('📝 Creating text template...');
        const txtContent = createFormattedTextTemplate(docDetails);
        const txtPath = path.join(DESKTOP_PATH, `${docDetails.type}_Template_${timestamp}.txt`);
        require('fs').writeFileSync(txtPath, txtContent);
        results.push(`📝 ${path.basename(txtPath)} - Text template (copy/paste ready)`);
        
        // 4. Create Word XML format (additional option)
        console.log('📋 Creating Word XML template...');
        const xmlContent = createWordXMLTemplate(docDetails);
        const xmlPath = path.join(DESKTOP_PATH, `${docDetails.type}_${timestamp}.xml`);
        require('fs').writeFileSync(xmlPath, xmlContent);
        results.push(`📋 ${path.basename(xmlPath)} - Word XML format`);

        return `📝 ADVANCED DOCUMENT TEMPLATES CREATED:

${docDetails.aiGenerated ? `🤖 AI-GENERATED CONTENT BY ${docDetails.generatedBy.toUpperCase()}` : '📝 STANDARD TEMPLATE CONTENT'}

📁 FILES CREATED ON DESKTOP:
${results.map(r => '• ' + r).join('\n')}

🎯 RECOMMENDED USAGE PRIORITY:
1. **Best Option**: Double-click the .rtf file (should open in Word with formatting)
2. **Alternative 1**: Right-click .rtf → "Open with" → Microsoft Word
3. **Alternative 2**: Open Word manually → File → Open → select .html file
4. **Manual Copy**: Use the .txt file for copy/paste into any application
5. **XML Option**: Advanced users can use the .xml file in Word

📄 DOCUMENT DETAILS:
• Type: ${docDetails.type}
• Title: ${docDetails.title}
• AI Generated: ${docDetails.aiGenerated ? 'Yes (' + docDetails.generatedBy + ')' : 'No (Template)'}
• Content Length: ${docDetails.content.length} characters
• Multiple formats created for maximum compatibility

💡 WHY MULTIPLE TEMPLATES:
• **RTF**: Best compatibility with Word, preserves formatting
• **HTML**: Modern format, can be opened by Word and browsers  
• **TXT**: Universal compatibility, copy/paste ready
• **XML**: Advanced Word format with rich formatting options

🔧 TROUBLESHOOTING:
If RTF files won't open in Word:
1. Right-click → "Open with" → Choose Microsoft Word
2. Or open Word first, then File → Open → select the file
3. Check if Word is properly installed and licensed
4. Try the HTML version as alternative

🤖 AI ENHANCEMENT NOTES:
${docDetails.aiGenerated ? 
`This content was intelligently generated using ${docDetails.generatedBy}, providing contextual and professional language appropriate for a ${docDetails.type.toLowerCase()}. The AI understood your request and created content that matches the intended purpose and tone.` : 
'Configure Claude API key to enable intelligent content generation that understands context and creates professional, purpose-specific content.'}`;

    } catch (error) {
        return `❌ Template creation failed: ${error.message}

This indicates severe system restrictions. Try:
1. Opening Word manually from Start menu
2. Creating document by hand using this content:

DOCUMENT TYPE: ${docDetails.type}
TITLE: ${docDetails.title || 'Untitled'}

CONTENT:
${docDetails.content}

MANUAL STEPS:
1. Copy the content above
2. Open Microsoft Word
3. Paste content (Ctrl+V)
4. Format as needed (fonts, headings, spacing)
5. Save with your preferred filename`;
    }
}

// Create advanced RTF with better formatting
function createAdvancedRTF(docDetails) {
    const rtfTitle = docDetails.title.replace(/[{}\\]/g, '').replace(/"/g, '');
    let rtfContent = docDetails.content
        .replace(/[{}\\]/g, '')
        .replace(/"/g, '')
        .replace(/\[HEADING\]/g, '')
        .replace(/\[\/HEADING\]/g, '')
        .replace(/\n\n/g, '\\par\\par ')
        .replace(/\n/g, '\\par ')
        .replace(/\[PARAGRAPH BREAK\]/g, '\\par\\par ');
    
    return `{\\rtf1\\ansi\\ansicpg1252\\deff0 
{\\fonttbl 
{\\f0\\froman\\fcharset0 Times New Roman;}
{\\f1\\fswiss\\fcharset0 Arial;}
{\\f2\\fmodern\\fcharset0 Courier New;}
}
{\\colortbl ;\\red0\\green0\\blue0;\\red0\\green0\\blue255;\\red128\\green128\\blue128;}
\\f0\\fs24\\cf1
${rtfTitle ? `{\\pard\\qc\\f1\\fs32\\b ${rtfTitle}\\b0\\fs24\\par\\par}` : ''}
{\\pard\\ql ${rtfContent}\\par}
${docDetails.aiGenerated ? `\\par\\par{\\pard\\ql\\fs18\\i\\cf3 ---\\par Generated by ${docDetails.generatedBy} on ${new Date().toLocaleDateString()} at ${new Date().toLocaleTimeString()}\\i0\\fs24\\cf1\\par}` : ''}
}`;
}

// Create Word-compatible HTML
function createWordCompatibleHTML(docDetails) {
    return `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head>
    <meta charset="UTF-8">
    <meta name="ProgId" content="Word.Document">
    <meta name="Generator" content="Microsoft Word">
    <meta name="Originator" content="Microsoft Word">
    <title>${docDetails.title || 'Document'}</title>
    <!--[if gte mso 9]>
    <xml>
        <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>90</w:Zoom>
            <w:DoNotPromptForConvert/>
            <w:DoNotShowRevisions/>
            <w:DoNotPrintRevisions/>
            <w:DisplayBackgroundShape/>
            <w:DoNotDisplayPageBoundaries/>
            <w:BackgroundShape>
                <v:background id="_x0000_s1025">
                    <v:fill>
                        <v:fill color="white" opacity="1"/>
                    </v:fill>
                </v:background>
            </w:BackgroundShape>
            <w:DocumentKind>DocumentNotSpecified</w:DocumentKind>
            <w:DrawingGridVerticalSpacing>18 pt</w:DrawingGridVerticalSpacing>
            <w:DisplayBackgroundShape/>
            <w:ValidateAgainstSchemas/>
            <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
            <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
            <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
            <w:Compatibility>
                <w:BreakWrappedTables/>
                <w:SnapToGridInCell/>
                <w:WrapTextWithPunct/>
                <w:UseAsianBreakRules/>
            </w:Compatibility>
        </w:WordDocument>
    </xml>
    <![endif]-->
    <style>
        body { 
            font-family: 'Times New Roman', serif; 
            font-size: 12pt; 
            line-height: 1.5; 
            margin: 1in; 
            color: black;
            background-color: white;
        }
        h1 { 
            font-family: Arial, sans-serif;
            font-size: 18pt; 
            font-weight: bold; 
            text-align: center; 
            margin-bottom: 24pt; 
            color: black;
        }
        p { 
            margin-bottom: 12pt; 
            text-align: justify;
        }
        .footer { 
            font-size: 10pt; 
            font-style: italic; 
            color: #808080; 
            border-top: 1px solid #cccccc; 
            padding-top: 12pt; 
            margin-top: 24pt; 
        }
        .heading {
            font-weight: bold;
            font-size: 14pt;
            margin-top: 12pt;
            margin-bottom: 6pt;
        }
    </style>
</head>
<body>
    ${docDetails.title ? `<h1>${docDetails.title}</h1>` : ''}
    <div>
        ${docDetails.content
            .replace(/\[HEADING\](.*?)\[\/HEADING\]/g, '<p class="heading">$1</p>')
            .replace(/\n\n/g, '</p><p>')
            .replace(/\n/g, '<br>')
            .replace(/^/, '<p>')
            .replace(/$/, '</p>')
        }
    </div>
    ${docDetails.aiGenerated ? `<div class="footer">
        <hr>
        Generated by ${docDetails.generatedBy} on ${new Date().toLocaleString()}<br>
        Document Type: ${docDetails.type}<br>
        Content Length: ${docDetails.content.length} characters
    </div>` : ''}
</body>
</html>`;
}

// Create Word XML template
function createWordXMLTemplate(docDetails) {
    const xmlContent = docDetails.content
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;')
        .replace(/\n/g, '</w:t></w:r></w:p><w:p><w:r><w:t>');
    
    return `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        ${docDetails.title ? `
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="32"/>
                </w:rPr>
                <w:t>${docDetails.title}</w:t>
            </w:r>
        </w:p>
        <w:p/>
        ` : ''}
        <w:p>
            <w:r>
                <w:t>${xmlContent}</w:t>
            </w:r>
        </w:p>
        ${docDetails.aiGenerated ? `
        <w:p/>
        <w:p>
            <w:r>
                <w:rPr>
                    <w:i/>
                    <w:sz w:val="18"/>
                    <w:color w:val="808080"/>
                </w:rPr>
                <w:t>---</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:rPr>
                    <w:i/>
                    <w:sz w:val="18"/>
                    <w:color w:val="808080"/>
                </w:rPr>
                <w:t>Generated by ${docDetails.generatedBy} on ${new Date().toLocaleDateString()}</w:t>
            </w:r>
        </w:p>
        ` : ''}
    </w:body>
</w:document>`;
}

// Create formatted text template
function createFormattedTextTemplate(docDetails) {
    return `${docDetails.type.toUpperCase()} TEMPLATE
${'='.repeat(60)}

${docDetails.aiGenerated ? `🤖 AI-GENERATED BY ${docDetails.generatedBy.toUpperCase()}` : 'STANDARD TEMPLATE'}
Created: ${new Date().toLocaleString()}

${'='.repeat(60)}

${docDetails.title ? `${docDetails.title.toUpperCase()}\n\n` : ''}${docDetails.content}

${'='.repeat(60)}

📋 COPY-PASTE INSTRUCTIONS:
1. Select and copy the content above (from title to end of content)
2. Open Microsoft Word manually
3. Create a new document (Ctrl+N)
4. Paste content (Ctrl+V)
5. Format as needed (fonts, spacing, etc.)
6. Save with your preferred filename

💡 FORMATTING TIPS:
• Make title larger and bold (select title → Home → Font Size → 18pt → Bold)
• Use proper paragraph spacing (Home → Paragraph → Line Spacing)
• Add page breaks if needed (Insert → Page Break)
• Adjust margins if required (Layout → Margins)
• Apply heading styles for better structure (Home → Styles)

🎨 SUGGESTED FORMATTING:
• Title: Arial, 18pt, Bold, Center-aligned
• Body: Times New Roman, 12pt, Left-aligned
• Line Spacing: 1.5 or Double
• Margins: Normal (1" all sides)

${docDetails.aiGenerated ? `

🤖 AI GENERATION NOTES:
This content was intelligently generated using ${docDetails.generatedBy}, providing:
• Contextual understanding of your request
• Professional language appropriate for a ${docDetails.type.toLowerCase()}
• Proper structure and formatting hints
• Content that matches the intended purpose and tone

The AI analyzed your request and created content specifically tailored to your needs.
You can further customize this content to match your exact requirements.` : ''}

🔧 TROUBLESHOOTING:
If you encounter issues:
• Make sure Microsoft Word is installed and properly licensed
• Try opening Word first, then pasting the content
• Check if you have sufficient permissions to create files
• Consider using the RTF or HTML files instead for better formatting

📁 ADDITIONAL FILES:
Check your Desktop for other formats of this document:
• .rtf file (Rich Text Format - best for Word)
• .html file (Web format - can be opened in Word)
• .xml file (Word XML format - advanced option)`;
}

// Enhanced document request parsing
function parseDocumentRequest(message) {
    const msg = message.toLowerCase();
    let type = 'Document';
    let title = '';
    let content = '';

    if (msg.includes('letter')) {
        type = 'Letter';
        if (msg.includes('principal')) {
            title = 'Letter to Principal';
            content = `Date: ${new Date().toLocaleDateString()}

To,
The Principal,
[School/College Name]
[Address]

Subject: [Subject Line]

Dear Sir/Madam,

I am writing this letter to bring to your attention [state your purpose].

[PARAGRAPH BREAK]

[Main content of your letter - explain the situation, request, or information you want to convey. Be specific about dates, circumstances, and any relevant details.]

[PARAGRAPH BREAK]

I would be grateful for your kind consideration and prompt response to this matter.

Thank you for your time and attention.

[PARAGRAPH BREAK]

Yours sincerely,

[Your Name]
[Your Class/Position]
[Contact Information]`;
        } else if (msg.includes('complaint')) {
            title = 'Complaint Letter';
            content = `Date: ${new Date().toLocaleDateString()}

To,
[Recipient Name]
[Position]
[Organization/Address]

Subject: Complaint Regarding [Issue]

Dear Sir/Madam,

I am writing to formally complain about [describe the issue].

[PARAGRAPH BREAK]

[Detailed explanation of the problem, including:]
• Date and time of occurrence
• Location where it happened
• People involved
• Impact on you or others
• Any attempts made to resolve the issue

[PARAGRAPH BREAK]

I would appreciate your immediate attention to this matter and expect a prompt resolution within [timeframe].

I look forward to your response and a satisfactory solution.

[PARAGRAPH BREAK]

Yours sincerely,

[Your Name]
[Contact Information]`;
        } else {
            title = 'Formal Letter';
            content = `Date: ${new Date().toLocaleDateString()}

[Recipient Name]
[Recipient Title/Position]
[Organization]
[Address]

Dear [Recipient Name],

I hope this letter finds you in good health and spirits.

[PARAGRAPH BREAK]

[State your purpose clearly in the opening paragraph]

[PARAGRAPH BREAK]

[Main content - provide details, explanations, or requests. Use multiple paragraphs as needed to organize your thoughts clearly.]

[PARAGRAPH BREAK]

[Closing paragraph - summarize your request and indicate next steps]

Thank you for your time and consideration.

[PARAGRAPH BREAK]

Best regards,

[Your Name]
[Your Position/Designation]
[Your Contact Information]`;
        }
    } else if (msg.includes('essay')) {
        type = 'Essay';
        title = 'Essay';
        content = `[HEADING]Introduction[/HEADING]

[Write your introduction paragraph here. Start with a hook to grab the reader's attention, provide background information on your topic, and conclude with a clear thesis statement that outlines the main points you will discuss.]

[PARAGRAPH BREAK]

[HEADING]Body Paragraph 1[/HEADING]

[Present your first main point with supporting evidence. Include specific examples, statistics, quotes, or research to support your argument. Explain how this evidence supports your thesis.]

[PARAGRAPH BREAK]

[HEADING]Body Paragraph 2[/HEADING]

[Present your second main point with supporting evidence. Maintain a logical flow from your previous paragraph and continue building your argument with concrete examples and analysis.]

[PARAGRAPH BREAK]

[HEADING]Body Paragraph 3[/HEADING]

[Present your third main point with supporting evidence. This paragraph should strengthen your overall argument and prepare the reader for your conclusion.]

[PARAGRAPH BREAK]

[HEADING]Conclusion[/HEADING]

[Summarize your main points and restate your thesis in different words. Provide final thoughts, implications of your argument, or a call to action. Leave the reader with something meaningful to consider.]`;
    } else if (msg.includes('report')) {
        type = 'Report';
        title = 'Report';
        content = `[HEADING]Executive Summary[/HEADING]

[Provide a brief overview of the report's main findings, conclusions, and recommendations. This should be concise but comprehensive enough for readers who may only read this section.]

[PARAGRAPH BREAK]

[HEADING]Introduction[/HEADING]

[Explain the background, purpose, and scope of the report. Include the problem or question being addressed and the methodology used.]

[PARAGRAPH BREAK]

[HEADING]Methodology[/HEADING]

[Describe how the information was gathered, what research methods were used, data sources, and any limitations in the approach.]

[PARAGRAPH BREAK]

[HEADING]Findings[/HEADING]

[Present the main results and observations. Use clear headings, bullet points, or numbered lists to organize information. Include relevant data, statistics, or evidence.]

Key findings include:
• [Finding 1]
• [Finding 2]  
• [Finding 3]

[PARAGRAPH BREAK]

[HEADING]Analysis[/HEADING]

[Interpret the findings and explain their significance. Discuss patterns, trends, or relationships discovered in the data.]

[PARAGRAPH BREAK]

[HEADING]Recommendations[/HEADING]

[Provide specific, actionable recommendations based on the findings. Prioritize them and explain the rationale for each.]

1. [Recommendation 1 with explanation]
2. [Recommendation 2 with explanation]
3. [Recommendation 3 with explanation]

[PARAGRAPH BREAK]

[HEADING]Conclusion[/HEADING]

[Summarize the key points and emphasize the importance of implementing the recommendations.]`;
    } else if (msg.includes('resume') || msg.includes('cv')) {
        type = 'Resume';
        title = 'Professional Resume';
        content = `[Your Name]
[Your Address] | [Phone Number] | [Email Address] | [LinkedIn Profile]

[PARAGRAPH BREAK]

[HEADING]Professional Summary[/HEADING]

[2-3 sentences describing your professional background, key skills, and career objectives]

[PARAGRAPH BREAK]

[HEADING]Experience[/HEADING]

[Job Title] | [Company Name] | [Dates]
• [Key responsibility or achievement]
• [Key responsibility or achievement]  
• [Key responsibility or achievement]

[Job Title] | [Company Name] | [Dates]
• [Key responsibility or achievement]
• [Key responsibility or achievement]
• [Key responsibility or achievement]

[PARAGRAPH BREAK]

[HEADING]Education[/HEADING]

[Degree] | [Institution] | [Year]
[Relevant coursework, honors, or GPA if applicable]

[PARAGRAPH BREAK]

[HEADING]Skills[/HEADING]

• Technical Skills: [List relevant technical skills]
• Software: [List software proficiencies]
• Languages: [List languages and proficiency levels]

[PARAGRAPH BREAK]

[HEADING]Certifications[/HEADING]

• [Certification Name] | [Issuing Organization] | [Date]
• [Certification Name] | [Issuing Organization] | [Date]`;
    } else if (msg.includes('poem')) {
        type = 'Poem';
        title = 'Creative Poem';
        content = `[Title of Your Poem]

[PARAGRAPH BREAK]

[First stanza]
[Write your creative verses here]
[Express your thoughts and emotions]
[Use rhythm, rhyme, and imagery]

[PARAGRAPH BREAK]

[Second stanza]
[Continue your poetic expression]
[Build upon your opening theme]
[Use metaphor and symbolism]

[PARAGRAPH BREAK]

[Third stanza]
[Develop your central message]
[Create vivid sensory details]
[Maintain your poetic voice]

[PARAGRAPH BREAK]

[Fourth stanza]
[Build towards your conclusion]
[Bring resolution or reflection]
[Leave a lasting impression]

[PARAGRAPH BREAK]

[Additional stanzas as needed to complete your poetic expression]`;
    } else {
        type = 'Document';
        title = 'New Document';
        content = `[HEADING]Document Overview[/HEADING]

Document created on ${new Date().toLocaleDateString()}

[PARAGRAPH BREAK]

[HEADING]Purpose[/HEADING]

[State the purpose of this document and what it aims to accomplish]

[PARAGRAPH BREAK]

[HEADING]Main Content[/HEADING]

[Start writing your main content here. Organize your thoughts into clear paragraphs with logical flow.]

Key Points:
• [Point 1 with supporting details]
• [Point 2 with supporting details]
• [Point 3 with supporting details]

[PARAGRAPH BREAK]

[HEADING]Additional Information[/HEADING]

[Include any additional relevant information, background context, or supporting details]

[PARAGRAPH BREAK]

[HEADING]Conclusion[/HEADING]

[Summarize the main points and provide final thoughts or next steps]`;
    }

    return {
        type: type,
        title: title,
        content: content,
        aiGenerated: false,
        generatedBy: null
    };
}

// Fallback to basic Word handling (for compatibility)
async function handleWord(message) {
    console.log('📝 Using basic Word handling (fallback mode)...');
    return await createAdvancedDocumentTemplate(message, null);
}

module.exports = {
    handleSmartWord,
    handleWord,
    detectWordComprehensively,
    attemptWordAutomationEnhanced,
    createDocumentAndOpenWord,
    createAdvancedDocumentTemplate,
    parseDocumentRequest
};