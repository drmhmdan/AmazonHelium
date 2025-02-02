#NoEnv  ; Recommended for performance and avoiding unexpected behaviour.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetBatchLines -1  ; Ensures the script runs at maximum speed.
#SingleInstance Force  ; Prevents multiple script instances from running.
; #Include <JSON>


clickWhenFound(imagePath) {
    Loop {
        ; Search for the image on the screen
        ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *80 photoes/%imagePath%    
        ; Check if the image was found
        if (ErrorLevel = 0) { ; ErrorLevel = 0 means the image was found
            MouseClick, left, %foundX%, %foundY%
            return true
        } else if (ErrorLevel = 1) { ; ErrorLevel = 1 means the image was not found
            Sleep, 100
            continue
        } else { ; ErrorLevel = 2 means there was an error (e.g., file not found)
            ;MsgBox, Error: Could not search for image "%imagePath%"
            return false
        }
    }
}

clickOneOfTwo(imagePath,imagePath2) {
    Loop {
        ; Search for the image on the screen
        ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *80 photoes/%imagePath%   
        ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *80 photoes/%imagePath2% 
        ; Check if the image was found
        if (ErrorLevel = 0) { ; ErrorLevel = 0 means the image was found
            MouseClick, left, %foundX%, %foundY%
            return true
        } else if (ErrorLevel = 1) { ; ErrorLevel = 1 means the image was not found
            Sleep, 100
            continue
        } else { ; ErrorLevel = 2 means there was an error (e.g., file not found)
            ;MsgBox, Error: Could not search for image "%imagePath%"
            return false
        }

        ; ; Search for the image on the screen
        ; ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *80 photoes/%imagePath2%    
        ; ; Check if the image was found
        ; if (ErrorLevel = 0) { ; ErrorLevel = 0 means the image was found
        ;     MouseClick, left, %foundX%, %foundY%
        ;     return true
        ; } else if (ErrorLevel = 1) { ; ErrorLevel = 1 means the image was not found
        ;     Sleep, 100
        ;     continue
        ; } else { ; ErrorLevel = 2 means there was an error (e.g., file not found)
        ;     ;MsgBox, Error: Could not search for image "%imagePath%"
        ;     return false
        ; }
    }
}

waitTillFound(imagePath) {
    Loop {
        ; Search for the image on the screen
        ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *60 photoes/%imagePath%    
        ; Check if the image was found
        if (ErrorLevel = 0) { ; ErrorLevel = 0 means the image was found
            ; move mouse to  %foundX%, %foundY%
            MouseMove, %foundX%, %foundY%
            return true
        } else if (ErrorLevel = 1) { ; ErrorLevel = 1 means the image was not found
            Sleep, 100
            continue
        } else { ; ErrorLevel = 2 means there was an error (e.g., file not found)
            ;MsgBox, Error: Could not search for image "%imagePath%"
            return false
        }
    }
}


startFlow() {
    ; openTheBrowser()
    Sleep, 300
    ; Click on the restore
    ; clickWhenFound("restoreLastSearch.png")

    Loop, 200 ; Loops 20 times
    {
        Sleep, 300
        ToolTip, now starting loop %A_Index%
        clickWhenFound("exportData.png")
        clickWhenFound("downloadCSV.png")

        Sleep, 1000

        clickWhenFound("criteriaRemoveDownladbox.png")

        ; Send the CSV file to the server and return JSON excluded words
        ; sendPostRequest("downloadsCSV/" . mostRecentCsv())
        ToolTip, now NOT sending downloadsCSV %A_Index%

        ;;;;;;;;;; for the old code with exclusions
        ; ; Read the JSON response for excluded words
        ; excludedWordsNew := ReadJson("downloadsCSV/" . mostRecentJSON())
        
        ; ; Append the new excluded words to the existing list in file excludedWords.csv
        ; FileAppend, %excludedWordsNew%, excludedWords.csv

        ; ; Read the existing excluded words from the file in the currect directory excludedWords.csv
        ; FileRead, allexcludedWords, excludedWords.csv
        ; newVariableExcluded := ""  ; Initialize a new variable to store the result

        ; ; Loop through each line in the file
        ; Loop, parse, allexcludedWords, `n, `r
        ; {
        ;     if (A_LoopField != "")  ; Ignore empty lines
        ;     {
        ;         if (newVariableExcluded != "")  ; Add a comma if newVariable is not empty
        ;             newVariableExcluded .= ", "
        ;         newVariableExcluded .= A_LoopField  ; Append the current line
        ;     }
        ; }

        Sleep, 500
        clickWhenFound("editFilters.png")
        Sleep, 500
        ; send scroll down with mouse using the send function
        Send, {WheelDown 7}
        Sleep, 100
        Send, {WheelDown 2}
        ; clickWhenFound("excludeTitleKeywords.png")
        clickWhenFound("preis.png")
        Sleep 100
        Send, {Tab}
        incriment := 0.5
        startPreis := 50.0 + (A_Index * incriment)
        endPreis := startPreis + (incriment)
        ; MsgBox, %startPreis%, "   ", %endPreis%

        Send, %startPreis%
        Sleep, 100
        Send, {Tab}
        Sleep, 100
        Send, %endPreis%
        Sleep, 100

        ; ; Send the excluded words to the input field
        ; ; Set the clipboard to the text you want to paste
        ; Clipboard := newVariableExcluded
        ; ; Wait for the clipboard to update (optional safeguard)
        ; ClipWait, 1
        ; ; Use the paste function (Ctrl+V)
        ; Send, ^v
        ; Sleep, 100

        clickWhenFound("applyFilters.png")
        Sleep, 300

        ToolTip, %endPreis%
        ;now we need to click export data again
    }
    
}


sendPostRequest(fileCSV) {
    ; Path to the file to upload
    filePath := fileCSV ; Replace this with the actual file path
    ToolTip, now sending %filePath%
    ; URL to send the request
    url := "http://89.40.6.97:3300/upload-products"

    ; Read the file content
    fileContent := ""
    FileRead, fileContent, %filePath%

    ; Create boundary for multipart form-data
    boundary := "----WebKitFormBoundary" . Format("{:08x}", A_TickCount)

    ; Build the multipart body
    body := "--" boundary "`r`n"
        . "Content-Disposition: form-data; name=""file""; filename=""" . FileExist(filePath) . """`r`n"
        . "Content-Type: text/plain`r`n`r`n" ; Replace text/plain with the actual file MIME type if necessary
        . fileContent "`r`n"
        . "--" boundary "--`r`n"

    ; Create a WinHTTP COM object
    http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    ToolTip, now sending %http%
    ; Open a synchronous POST request
    http.Open("POST", url, false)

    ; Set the headers
    http.SetRequestHeader("Content-Type", "multipart/form-data; boundary=" boundary)
    http.SetRequestHeader("Content-Length", StrLen(body))
    ToolTip, now sending body %A_Index%
    ; Send the request
    http.Send(body)
    ToolTip, now body sent
    ; Save the response to a text file
    responseFile := fileCSV . "ExcludedWords.json"
    FileDelete, %responseFile% ; Delete the file if it already exists
    FileAppend, % http.ResponseText, %responseFile%

    ; Display a message box indicating completion
    Tooltip, The response has been saved to %responseFile%.


}


ReadJson(csvJSON) {
    ; Read the JSON file
    jsonFile := csvJSON

    ; Read the JSON file content
    FileRead, jsonContent, %jsonFile%

    ; Parse the JSON content
    parsedJSON := JSON.Load(jsonContent)

    ; Access fields
    message := parsedJSON["message"]
    keywords := parsedJSON["keywords"]

    ; Display extracted fields
    ;MsgBox, %keywords%
    Return keywords

}


sendPostGemini() {
    ; Define the API endpoint and key
    api_url := "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=AIzaSyDRLp5V09OWy47IZCSQ_PMEJUZ7jXxrmh4"

    ; Define the JSON payload
    json_payload := "
    (
    {
    ""contents"": [{
        ""parts"": [{""text"": ""in two sentences, Explain how AI works""}]
    }]
    }
    )"

    ; Create the HTTP request
    http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    http.Open("POST", api_url, false)
    http.SetRequestHeader("Content-Type", "application/json")
    http.Send(json_payload)

    ; Get the response
    response := http.ResponseText

    ; Show the response in a message box
    ;MsgBox, % response

}

startFlowStep2(){
    ;open the browser manually 
    ;https://www.amazon.de/s?k=(Keyword)

    ;read the txt file with the keywords line by line naes test.csv using uft-8 .Encoding


    FileRead, keywords, competition_keywords_results.csv ; Read the entire file غير اسم الفايل الي فيه الملفات ... لازم يكون فقط كيوورد ورا بعض
    ;loop through each keyword from keywords
    Loop, parse, keywords, `n, `r
    {
        Sleep 100
        Send, ^l
        Sleep 300
        ;send the keyword to the search bar in url plus ttps://www.amazon.de/s?k=
        SendInput, https://www.amazon.de/s?k=%A_LoopField%
        Sleep 100
        Send, {Enter}
        Sleep 100
        ;tooltip showing the keyword and loopcount
        ToolTip, %A_LoopField% %A_Index%
        Send, ^e
        clickWhenFound("amazonProduktRese.png")
        ; waitTillFound("export.png")
        ; clickWhenFound("suchVolumen.png")
        ; Send, {Shift down}{Down}{Down}{Shift up}
        ; MsgBox, Pauseff
        ; Sleep 100
        ; Clipboard := ""  ; Clear the clipboard to start fresh
        ; Send, ^c         ; Simulate Ctrl+C to copy the highlighted text
        ; ClipWait, 1      ; Wait for the clipboard to contain data
        ; Loop             ; Start a loop
        ; {
        ;     Sleep, 100   ; Wait 100ms before checking the clipboard
        ;     if (Clipboard != "")  ; Check if the clipboard has content
        ;         break     ; Exit the loop if clipboard contains something
        ; }
        ; Sleep 100
        ; ;save the Clipboard to a variable called suchVolumenClipboard
        ; suchVolumenClipboard := RegExReplace(Clipboard, "[\s/]", "") ; Remove all spaces, newlines, and "/"
        clickWhenFound("parentLevelSales.png")
        Sleep 2000 ; wait till sorted
        clickWhenFound("export.png")
        clickWhenFound("CSVDatei.png")
        ; clickOneOfTwo("saveAs.png","saveAs2.png") ;enable for edge
        clickWhenFound("fileName.png")
        Sleep 100
        SendInput, %A_LoopField%_%suchVolumenClipboard%.csv
        ; MsgBox, Pauseff
        Sleep 100
        Send, {Enter}
        Sleep 700
        ; Send, {F4}
        Sleep 100

    }





}


^k::startFlowStep2()  ; Ctrl+K to start the flow هاد الترجر الي ببلش الكود




;finishing
!r::Reload
Esc::Pause