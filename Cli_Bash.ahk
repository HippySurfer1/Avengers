; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


F3::
Macro1:
/*
SendRaw, Run the daemon. Wait for sync.
*/
Run, C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe, C:\Program Files\Monero GUI Wallet
Sleep, 300
Send, .\monerod.exe --data-dir E:\BitMonero {Enter}
MsgBox, 0, , Please click OK when the wallet is synced
x := 1
/*
SendRaw, Wallet loop
*/
Loop
{
    FileReadLine, Wallet, D:\Monero\wallets\Wallets.txt, %X%
    x := x+1
    Run, C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe, C:\Program Files\Monero GUI Wallet
    Sleep, 300
    Send, .\monero-wallet-cli --wallet-file=D:\Monero\wallets\%Wallet% --password Ra1jlt01{Enter}
    Sleep, 5000
    Send, refresh {Enter}
    Loop
    {
        CoordMode, Pixel, Window
        PixelSearch, FoundX, FoundY, 44, 532, 44, 532, 0xFFFF00, 0, Fast RGB
        Sleep, 3
    }
    Until ErrorLevel = 0
    If (ErrorLevel = 0)
    {
        /*
        SendRaw, Export and exit powershell
        */
    }
    Send, export_transfers {Enter}
    Sleep, 1000
    Send, ^c
    Sleep, 300
    Send, exit{Enter}
    Sleep, 300
    FileMove, C:\Program Files\Monero GUI Wallet\output0.csv, D:\MoneroData\%Wallet%\txs.csv
    Run, C:\Program Files\Git\git-bash.exe, D:\MoneroData\%Wallet%
    /*
    SendRaw, Find new transactions
    */
    Sleep, 3000
    Send, comm  -1 -3 --nocheck-order  txs-old.csv txs.csv > diff.csv{Enter} 
    Sleep, 3000
    FileMove, D:\MoneroData\%Wallet%\diff.csv, D:\MoneroData\Common\diff.csv
    /*
    SendRaw, 3 day backup
    */
    Sleep, 300
    Send, mv txs-old-old.csv txs-old-old-old.csv{Enter}
    Sleep, 3000
    Send, mv txs-old.csv txs-old-old.csv{Enter}
    Sleep, 3000
    Send, mv txs.csv txs-old.csv{Enter}
    Sleep, 3000
    Send, exit{Enter}
    Run, D:\Excel\Book1.xlsm
    WinWait, Microsoft Excel - Book1.xlsm
    Sleep, 333
    /*
    SendRaw, Load Excel and run Excel VB import
    */
    Sleep, 300
    If (!IsObject(XL))
        XL := ComObjCreate("Excel.Application")
    XL := ComObjActive("Excel.Application")
    Send, (Alt down}{F8down}
    SendRaw, {Alt up}{F8 up}
    Sleep, 300
    /*
    SendRaw, Excel to QB loop
    */
    Send, ^r
    XL.Range("D1").Select
    Loop
    {
        Inout := XL.ActiveCell.Text
        /*
        SendRaw, Inward to QB
        */
        IfInString, Inout, in
        {
            XL.ActiveCell.Offset(0, 11).Select
            Name := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, -3).Select
            Amount := XL.ActiveCell.Value
            StringTrimRight, TrimAmount, Amount, 4
            WinActivate, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
            Sleep, 333
            WinMenuSelectItem, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant), , 6&, 2&
            /*
            Send, {Alt Down}{u Down}
            Sleep, 100
            Send, {Alt Up}{u Up}
            SendRaw, s
            */
            Sleep, 3000
            Send, %Name%
            MsgBox, 0, , Stop
            Sleep, 100
            /*
            Send, {Enter}
            */
            Sleep, 100
            Send, {Tab}
            SendRaw, Intuit
            Sleep, 100
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            SendRaw, Don
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, %TrimAmount%
            Send, {Tab}
            Sleep, 100
            Send, {Enter}
            Send, {Alt Down}{a Down}
            Send, {Alt Up}{a Up}
            XL.ActiveCell.Offset(1, -8).Select
            /*
            SendRaw, Outwad to QB
            */
        }
        Else IfInString, Inout, out
        {
            XL.ActiveCell.Offset(0, 4).Select
            Name := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 4).Select
            Amount := XL.ActiveCell.Value
            StringTrimLeft, TrimName, Name, 30
            StringTrimRight, TrimAmount, Amount, 4
            WinActivate, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
            Sleep, 333
            Send, ^w
            Sleep, 2000
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, %TrimName%{Tab}
            Sleep, 100
            Send, %TrimAmount%{Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, {Tab}
            Sleep, 100
            Send, Test Payment{Tab}
            Sleep, 100
            Send, %TrimAmount%{Tab}
            Sleep, 300
            Send, {Alt Down}{a Down}
            Send, {a Up}{Alt Up}
            XL := ComObjActive("Excel.Application")
            XL.ActiveCell.Offset(1, -8).Select
        }
    }
}
}
}
Until, %Wallet% := ""
Return

