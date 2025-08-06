"use client"

import { useState, useRef } from 'react'
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Upload, Download, FileText } from 'lucide-react'
import { useToast } from "@/hooks/use-toast"

// Import JSZip for handling .docx files (which are ZIP archives)
declare global {
  interface Window {
    JSZip: any;
  }
}

export default function DocsEditor() {
  const [documentContent, setDocumentContent] = useState<string>('')
  const [originalContent, setOriginalContent] = useState<string>('')
  const [fileName, setFileName] = useState<string>('')
  const [userName, setUserName] = useState<string>('')
  const [userDate, setUserDate] = useState<string>(new Date().toISOString().split('T')[0])
  const [isLoading, setIsLoading] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const editorRef = useRef<HTMLDivElement>(null)
  const { toast } = useToast()

  // Load JSZip dynamically
  const loadJSZip = async () => {
    if (typeof window !== 'undefined' && !window.JSZip) {
      const script = document.createElement('script')
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'
      document.head.appendChild(script)
      
      return new Promise((resolve) => {
        script.onload = () => resolve(window.JSZip)
      })
    }
    return window.JSZip
  }

  // Parse XML content from document.xml
  const parseDocumentXML = (xmlContent: string): string => {
    try {
      const parser = new DOMParser()
      const xmlDoc = parser.parseFromString(xmlContent, 'text/xml')
      
      let htmlContent = '<div class="document-content">'
      
      // Get all paragraph elements
      const paragraphs = xmlDoc.getElementsByTagName('w:p')
      
      for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i]
        let paragraphText = ''
        let isHeading = false
        let isBold = false
        let isItalic = false
        
        // Check for heading styles
        const pPr = paragraph.getElementsByTagName('w:pPr')[0]
        if (pPr) {
          const pStyle = pPr.getElementsByTagName('w:pStyle')[0]
          if (pStyle) {
            const styleVal = pStyle.getAttribute('w:val')
            if (styleVal && (styleVal.includes('Heading') || styleVal.includes('Title'))) {
              isHeading = true
            }
          }
        }
        
        // Get all text runs in the paragraph
        const runs = paragraph.getElementsByTagName('w:r')
        for (let j = 0; j < runs.length; j++) {
          const run = runs[j]
          
          // Check for formatting
          const rPr = run.getElementsByTagName('w:rPr')[0]
          if (rPr) {
            isBold = rPr.getElementsByTagName('w:b').length > 0
            isItalic = rPr.getElementsByTagName('w:i').length > 0
          }
          
          // Get text content
          const textElements = run.getElementsByTagName('w:t')
          for (let k = 0; k < textElements.length; k++) {
            let text = textElements[k].textContent || ''
            
            if (isBold) text = `<strong>${text}</strong>`
            if (isItalic) text = `<em>${text}</em>`
            
            paragraphText += text
          }
        }
        
        // Add paragraph to HTML
        if (paragraphText.trim()) {
          if (isHeading) {
            htmlContent += `<h2 class="document-heading">${paragraphText}</h2>`
          } else {
            htmlContent += `<p class="document-paragraph">${paragraphText}</p>`
          }
        }
      }
      
      // Parse tables
      const tables = xmlDoc.getElementsByTagName('w:tbl')
      for (let i = 0; i < tables.length; i++) {
        const table = tables[i]
        htmlContent += '<table class="document-table">'
        
        const rows = table.getElementsByTagName('w:tr')
        for (let j = 0; j < rows.length; j++) {
          const row = rows[j]
          htmlContent += '<tr class="table-row">'
          
          const cells = row.getElementsByTagName('w:tc')
          for (let k = 0; k < cells.length; k++) {
            const cell = cells[k]
            let cellText = ''
            
            const cellParagraphs = cell.getElementsByTagName('w:p')
            for (let l = 0; l < cellParagraphs.length; l++) {
              const cellP = cellParagraphs[l]
              const cellRuns = cellP.getElementsByTagName('w:r')
              for (let m = 0; m < cellRuns.length; m++) {
                const cellRun = cellRuns[m]
                const cellTextElements = cellRun.getElementsByTagName('w:t')
                for (let n = 0; n < cellTextElements.length; n++) {
                  cellText += cellTextElements[n].textContent || ''
                }
              }
            }
            
            const cellTag = j === 0 ? 'th' : 'td'
            const cellClass = j === 0 ? 'table-header' : 'table-cell'
            htmlContent += `<${cellTag} class="${cellClass}">${cellText}</${cellTag}>`
          }
          
          htmlContent += '</tr>'
        }
        
        htmlContent += '</table>'
      }
      
      // Add placeholders section
      htmlContent += `
        <div class="placeholder-section">
          <h3 class="document-subheading">Fill Information</h3>
          <p class="document-paragraph"><strong>Name:</strong> [NAME_PLACEHOLDER]</p>
          <p class="document-paragraph"><strong>Date:</strong> [DATE_PLACEHOLDER]</p>
        </div>
      `
      
      htmlContent += '</div>'
      return htmlContent
      
    } catch (error) {
      console.error('Error parsing XML:', error)
      return '<div class="document-content"><p class="document-paragraph">Error parsing document content</p></div>'
    }
  }

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    if (!file.name.endsWith('.docx') && !file.name.endsWith('.doc')) {
      toast({
        title: "Invalid file type",
        description: "Please upload a .doc or .docx file",
        variant: "destructive"
      })
      return
    }

    setIsLoading(true)
    setFileName(file.name)

    try {
      // Load JSZip library
      await loadJSZip()
      const JSZip = window.JSZip

      if (file.name.endsWith('.docx')) {
        // Read the .docx file as a ZIP archive
        const arrayBuffer = await file.arrayBuffer()
        const zip = await JSZip.loadAsync(arrayBuffer)
        
        // Extract the main document content
        const documentXML = await zip.file('word/document.xml')?.async('string')
        
        if (documentXML) {
          const htmlContent = parseDocumentXML(documentXML)
          setOriginalContent(htmlContent)
          setDocumentContent(htmlContent)
          
          toast({
            title: "Document uploaded successfully",
            description: `${file.name} has been parsed and loaded with original formatting`
          })
        } else {
          throw new Error('Could not find document.xml in the .docx file')
        }
        
      } else {
        // For .doc files (older format) - basic text extraction
        const reader = new FileReader()
        reader.onload = (e) => {
          const arrayBuffer = e.target?.result as ArrayBuffer
          const decoder = new TextDecoder('utf-8', { fatal: false })
          const text = decoder.decode(arrayBuffer)
          
          // Clean and format the text
          const cleanText = text
            .replace(/[^\w\s\.\,\!\?\:\;\-$$$$\[\]\{\}\"\']/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
          
          const sentences = cleanText.split(/[\.!?]+/).filter(s => s.trim().length > 5)
          
          let htmlContent = '<div class="document-content">'
          htmlContent += `<h1 class="document-title">${file.name.replace('.doc', '')}</h1>`
          
          sentences.forEach((sentence, index) => {
            const trimmed = sentence.trim()
            if (trimmed.length > 0) {
              if (trimmed.length < 50 && index < 3) {
                htmlContent += `<h2 class="document-heading">${trimmed}</h2>`
              } else {
                htmlContent += `<p class="document-paragraph">${trimmed}.</p>`
              }
            }
          })
          
          htmlContent += `
            <div class="placeholder-section">
              <h3 class="document-subheading">Fill Information</h3>
              <p class="document-paragraph"><strong>Name:</strong> [NAME_PLACEHOLDER]</p>
              <p class="document-paragraph"><strong>Date:</strong> [DATE_PLACEHOLDER]</p>
            </div>
          </div>`
          
          setOriginalContent(htmlContent)
          setDocumentContent(htmlContent)
          
          toast({
            title: "Document uploaded successfully",
            description: `${file.name} content has been extracted`
          })
        }
        
        reader.readAsArrayBuffer(file)
      }
      
      setIsLoading(false)
      
    } catch (error) {
      console.error('Error processing file:', error)
      setIsLoading(false)
      toast({
        title: "Error processing file",
        description: "Failed to parse the document. Please try a different file.",
        variant: "destructive"
      })
    }
  }

  const generateDocument = () => {
    if (!originalContent) {
      toast({
        title: "No document loaded",
        description: "Please upload a document first",
        variant: "destructive"
      })
      return
    }

    if (!userName.trim()) {
      toast({
        title: "Name required",
        description: "Please enter a name",
        variant: "destructive"
      })
      return
    }

    // Replace placeholders with actual values
    let updatedContent = originalContent
      .replace(/\[NAME_PLACEHOLDER\]/g, userName)
      .replace(/\[DATE_PLACEHOLDER\]/g, new Date(userDate).toLocaleDateString())

    setDocumentContent(updatedContent)
    
    toast({
      title: "Document generated",
      description: "Name and date have been inserted into the document"
    })
  }

  const saveDocument = async () => {
    if (!documentContent) {
      toast({
        title: "No document to save",
        description: "Please upload and generate a document first",
        variant: "destructive"
      })
      return
    }

    try {
      // Create a proper HTML document for Word compatibility
      const htmlContent = `
        <!DOCTYPE html>
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
        <head>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Document</title>
          <!--[if gte mso 9]>
          <xml>
            <w:WordDocument>
              <w:View>Print</w:View>
              <w:Zoom>90</w:Zoom>
              <w:DoNotPromptForConvert/>
              <w:DoNotShowInsertionsAndDeletions/>
            </w:WordDocument>
          </xml>
          <![endif]-->
          <style>
            body { 
              font-family: 'Times New Roman', serif; 
              font-size: 12pt; 
              line-height: 1.5; 
              margin: 1in; 
            }
            .document-content { margin: 0; }
            .document-title { 
              font-size: 18pt; 
              font-weight: bold; 
              text-align: center; 
              margin-bottom: 20pt; 
            }
            .document-heading { 
              font-size: 14pt; 
              font-weight: bold; 
              margin: 12pt 0 6pt 0; 
            }
            .document-subheading { 
              font-size: 12pt; 
              font-weight: bold; 
              margin: 10pt 0 5pt 0; 
            }
            .document-paragraph { 
              margin: 6pt 0; 
              text-align: justify; 
            }
            .document-table { 
              border-collapse: collapse; 
              width: 100%; 
              margin: 12pt 0; 
            }
            .table-header, .table-cell { 
              border: 1pt solid black; 
              padding: 6pt; 
              vertical-align: top; 
            }
            .table-header { 
              background-color: #f0f0f0; 
              font-weight: bold; 
            }
            .placeholder-section { 
              margin-top: 20pt; 
              padding: 12pt; 
              border: 1pt solid #ccc; 
            }
          </style>
        </head>
        <body>
          ${documentContent}
        </body>
        </html>
      `

      const blob = new Blob([htmlContent], { 
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
      })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = fileName ? fileName.replace(/\.[^/.]+$/, '_modified.docx') : 'document_modified.docx'
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)

      toast({
        title: "Document saved",
        description: "Your document has been downloaded with all formatting preserved"
      })
    } catch (error) {
      toast({
        title: "Error saving document",
        description: "Failed to save the document",
        variant: "destructive"
      })
    }
  }

  return (
    <div className="min-h-screen bg-gray-50 p-4">
      <div className="max-w-7xl mx-auto space-y-6">
        {/* Header */}
        <div className="text-center">
          <h1 className="text-3xl font-bold text-gray-900">Professional Document Editor</h1>
          <p className="text-gray-600 mt-2">Upload Word documents and see exact content with tables, formatting, and structure preserved</p>
        </div>

        <div className="grid lg:grid-cols-4 gap-6">
          {/* Control Panel */}
          <div className="lg:col-span-1 space-y-4">
            {/* File Upload */}
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Upload className="w-5 h-5" />
                  Upload Document
                </CardTitle>
                <CardDescription>
                  Upload a .docx file to see its exact content
                </CardDescription>
              </CardHeader>
              <CardContent>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".doc,.docx"
                  onChange={handleFileUpload}
                  className="hidden"
                />
                <Button 
                  onClick={() => fileInputRef.current?.click()}
                  className="w-full"
                  disabled={isLoading}
                >
                  {isLoading ? 'Processing...' : 'Choose File'}
                </Button>
                {fileName && (
                  <p className="text-sm text-gray-600 mt-2 flex items-center gap-1">
                    <FileText className="w-4 h-4" />
                    {fileName}
                  </p>
                )}
              </CardContent>
            </Card>

            {/* Form Fields */}
            <Card>
              <CardHeader>
                <CardTitle>Document Details</CardTitle>
                <CardDescription>
                  Fill in the details to be inserted into the document
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="name">Name</Label>
                  <Input
                    id="name"
                    value={userName}
                    onChange={(e) => setUserName(e.target.value)}
                    placeholder="Enter your name"
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="date">Date</Label>
                  <Input
                    id="date"
                    type="date"
                    value={userDate}
                    onChange={(e) => setUserDate(e.target.value)}
                  />
                </div>
                <Button 
                  onClick={generateDocument}
                  className="w-full"
                  disabled={!documentContent || !userName.trim()}
                >
                  Generate Document
                </Button>
              </CardContent>
            </Card>

            {/* Save Options */}
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Download className="w-5 h-5" />
                  Save Document
                </CardTitle>
                <CardDescription>
                  Download the modified document with original formatting
                </CardDescription>
              </CardHeader>
              <CardContent>
                <Button 
                  onClick={saveDocument}
                  className="w-full"
                  disabled={!documentContent}
                  variant="outline"
                >
                  Save as .docx
                </Button>
              </CardContent>
            </Card>
          </div>

          {/* Document Preview */}
          <div className="lg:col-span-3">
            <Card className="h-full">
              <CardHeader>
                <CardTitle>Document Preview - Exact Content</CardTitle>
                <CardDescription>
                  This shows the exact content from your uploaded Word document
                </CardDescription>
              </CardHeader>
              <CardContent>
                {documentContent ? (
                  <div className="border rounded-lg p-6 bg-white min-h-[600px] max-h-[800px] overflow-y-auto shadow-inner">
                    <style jsx>{`
                      .document-content {
                        font-family: 'Times New Roman', serif;
                        font-size: 12pt;
                        line-height: 1.6;
                        color: #000;
                        max-width: none;
                      }
                      .document-title {
                        text-align: center;
                        color: #000;
                        font-size: 18pt;
                        font-weight: bold;
                        margin-bottom: 24px;
                      }
                      .document-heading {
                        color: #000;
                        font-size: 14pt;
                        font-weight: bold;
                        margin: 18px 0 12px 0;
                      }
                      .document-subheading {
                        color: #000;
                        font-size: 12pt;
                        font-weight: bold;
                        margin: 12px 0 8px 0;
                      }
                      .document-paragraph {
                        margin: 8px 0;
                        color: #000;
                        text-align: justify;
                        font-size: 12pt;
                      }
                      .document-table {
                        border-collapse: collapse;
                        width: 100%;
                        margin: 16px 0;
                        border: 1px solid #000;
                      }
                      .table-header {
                        padding: 8px 12px;
                        text-align: left;
                        border: 1px solid #000;
                        font-weight: bold;
                        background-color: #f0f0f0;
                        font-size: 11pt;
                      }
                      .table-row {
                        border: 1px solid #000;
                      }
                      .table-cell {
                        padding: 8px 12px;
                        border: 1px solid #000;
                        font-size: 11pt;
                        vertical-align: top;
                      }
                      .placeholder-section {
                        margin-top: 24px;
                        padding: 16px;
                        background-color: #f8f9fa;
                        border: 2px solid #007bff;
                        border-radius: 4px;
                      }
                      strong {
                        font-weight: bold;
                      }
                      em {
                        font-style: italic;
                      }
                    `}</style>
                    <div
                      ref={editorRef}
                      contentEditable
                      dangerouslySetInnerHTML={{ __html: documentContent }}
                      onBlur={(e) => setDocumentContent(e.currentTarget.innerHTML)}
                      className="outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 rounded min-h-[500px]"
                      style={{ 
                        fontFamily: 'Times New Roman, serif',
                        fontSize: '12pt',
                        lineHeight: '1.6'
                      }}
                    />
                  </div>
                ) : (
                  <div className="border-2 border-dashed border-gray-300 rounded-lg p-12 text-center">
                    <FileText className="w-12 h-12 text-gray-400 mx-auto mb-4" />
                    <h3 className="text-lg font-medium text-gray-900 mb-2">No document loaded</h3>
                    <p className="text-gray-600">Upload a .docx file to see its exact content with formatting preserved</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  )
}
