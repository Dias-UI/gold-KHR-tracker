#NoEnv
#Warn
#SingleInstance Force
SendMode Input

; Auto-execute section
#Persistent
SetWorkingDir %A_ScriptDir%

; Create GUI first
Gui, New                                              ; Initialize a new GUI window
Gui, +LastFound +AlwaysOnTop -Caption +ToolWindow    ; Make borderless, always on top window
Gui, Color, FFFFFF                                   ; Set background color to white
WinSet, TransColor, FFFFFF 220                       ; Make background semi-transparent
Gui, Font, s12 c00FFFF, Consolas                    ; Set cyan colored text, 12pt Consolas font for alignment
Gui, Add, Text, vGoldText w300 h300 Left, Initializing...  ; Increased height to 300

; Get screen dimensions and calculate position
SysGet, MonitorWorkArea, MonitorWorkArea             ; Get usable screen area excluding taskbar
GuiWidth := 300                                      ; Width of the window
GuiHeight := 500                                     ; Increased height to 500
xPos := MonitorWorkAreaRight - GuiWidth - 50         ; Position from right edge (reduced margin)
yPos := MonitorWorkAreaBottom - GuiHeight + 80      ; Moved window slightly lower

; Show GUI at calculated position
Gui, Show, x%xPos% y%yPos% w%GuiWidth% h%GuiHeight%

; Start update timer
SetTimer, UpdateData, 60000
GoSub, UpdateData

Return

; Function to format numbers with commas
FormatWithCommas(number) {
    return RegExReplace(number, "(\d)(?=(\d{3})+(\.|$))", "$1,")
}

; Function to right pad spaces for alignment
RightPad(number) {
    Loop, % 12 - StrLen(number)
    {
        number := " " . number
    }
    return number
}

UpdateData:
try {
    ; Get gold prices
    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", "https://www.livepriceofgold.com/cambodia-gold-price.html", true)
    whr.Send()
    whr.WaitForResponse()
    html := whr.ResponseText

    ; Get NBC rate
    nbc_whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    nbc_whr.Open("GET", "https://www.nbc.gov.kh/english/economic_research/exchange_rate.php", true)
    nbc_whr.Send()
    nbc_whr.WaitForResponse()
    nbc_html := nbc_whr.ResponseText

    ; Get Tax department rate
    tax_whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    tax_whr.Open("GET", "https://www.tax.gov.kh/en/exchange-rate", true)
    tax_whr.Send()
    tax_whr.WaitForResponse()
    tax_html := tax_whr.ResponseText

    ; Extract rates using RegEx
    RegExMatch(nbc_html, "Official Exchange Rate : <font color=""#FF3300"">(\d+)</font>", nbcMatch)
    RegExMatch(tax_html, "<span class=""moul"">(\d+)\s+Riel / USD</span>", taxMatch)

    nbc_rate := nbcMatch1 . ".00"
    tax_rate := taxMatch1 . ".00"

    ; Simplified RegEx pattern
    RegExMatch(html, "<span[^>]*id=""ozvalue""[^>]*>([\d.,]+)</span>", goldMatch)
    spot_oz := StrReplace(goldMatch1, ",", "")
    
    ; Calculate all unit conversions
    grams_per_oz := 31.1035
    grams_per_chi := 3.75
    
    price_chi := Round((spot_oz / grams_per_oz) * grams_per_chi, 2)
    price_domleng := Round(price_chi * 1.5, 2)
    price_hun := Round(price_chi * 0.1, 2)
    price_ly := Round(price_chi * 0.01, 2)
    price_gram := Round(spot_oz / grams_per_oz, 2)
    price_kilo := Round((spot_oz / grams_per_oz) * 1000, 2)

    ; Get exchange rates
    RegExMatch(html, "USD/KHR[\s\S]*?<td[^>]*?>(\d{1,3}(?:,\d{3})*\.\d{2})</td>", fxMatch)
    usd_khr := fxMatch1
    
    ; Get live exchange rate
    RegExMatch(html, "<span[^>]*data-price=""USDKHR""[^>]*>(\d{1,3}(?:,\d{3})*\.\d{2})</span>", liveMatch)
    live_rate := liveMatch1

    ; Format numbers with commas
    price_chi := FormatWithCommas(price_chi)
    price_hun := FormatWithCommas(price_hun)
    price_ly := FormatWithCommas(price_ly)
    price_domleng := FormatWithCommas(price_domleng)
    spot_oz := FormatWithCommas(spot_oz)
    price_gram := FormatWithCommas(price_gram)
    price_kilo := FormatWithCommas(price_kilo)
    usd_khr := FormatWithCommas(usd_khr)
    live_rate := FormatWithCommas(live_rate)
    nbc_rate := FormatWithCommas(nbc_rate)
    tax_rate := FormatWithCommas(tax_rate)

    displayText := "Gold Prices:`n"
    . "Chi:     $" . RightPad(price_chi) . "`n"
    . "Hun:     $" . RightPad(price_hun) . "`n"
    . "Ly:      $" . RightPad(price_ly) . "`n"
    . "Domleng: $" . RightPad(price_domleng) . "`n"
    . "oz:      $" . RightPad(spot_oz) . "`n"
    . "Gram:    $" . RightPad(price_gram) . "`n"
    . "Kilo:    $" . RightPad(price_kilo) . "`n`n"
    . "Exchange Rates:`n"
    . "Market:   " . RightPad(usd_khr) . "`n"
    . "NBC:      " . RightPad(nbc_rate) . "`n"
    . "Tax:      " . RightPad(tax_rate)

    GuiControl,, GoldText, %displayText%
} catch e {
    GuiControl,, GoldText, Error fetching data
    SetTimer, UpdateData, -5000  ; Retry after 5 seconds
}
return

!Esc::ExitApp
!r::Reload

^+LButton::
PostMessage, 0xA1, 2
return
