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
Run, C:/Program Files (x86)/Intuit/QuickBooks Enterprise Solutions 18.0/QBW32EnterpriseAccountant.exe, C:/Program Files (x86)/Intuit/QuickBooks Enterprise Solutions 18.0
Sleep, 25000
WinWait, AvengersAustralia  - Intuit QuickBooks Enterprise Solutions: Accountant 18.0, , 35000
Sleep, 333
Sleep, 300
WinMenuSelectItem, AvengersAustralia  - Intuit QuickBooks Enterprise Solutions: Accountant 18.0 ahk_class MauiFrame ahk_exe QBW32.EXE, , File, Toggle to Another Edition
Sleep, 500
Send, {Enter}
Sleep, 2000
WinActivate, Select QuickBooks Desktop Industry-Specific Edition
Sleep, 333
Send, {Down}{Down}{Down}
Send, {Space}
Send, {Enter}
Send, {Enter}
WinWait, AvengersAustralia  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
Sleep, 333
Run, C:/Windows/SysWOW64/WindowsPowerShell/v1.0/powershell.exe, C:/Monero
Sleep, 1000
Send, ./monerod.exe --data-dir E:/BitMonero {Enter}
Sleep, 300
MsgBox, 262144, , Please click OK when the daemon is synced
x := 0
Inout := "1"
Outer:
Loop
{
    x := x+"1"
    FileReadLine, Wallet, D:\Monero\Wallets.txt, %x%
    Run, C:/Windows/SysWOW64/WindowsPowerShell/v1.0/powershell.exe, C:/Monero
    Sleep, 500
    Send, ./monero-wallet-cli --wallet-file=D:/Monero/wallets/%Wallet% --password Ra1jlt01
    Sleep, 500
    Send, {Enter}
    Sleep, 500
    Send, refresh {Enter}
    MsgBox, 262144, , Please click OK when the wallet is synced and QuickBooks is loaded
    Send, export_transfers {Enter}
    Sleep, 500
    Send, exit{Enter}
    Sleep, 500
    Send, exit{Enter}
    FileMove, C:/Monero/output0.csv, D:/Monero/wallets/%Wallet%1/txs.csv
    Run, C:/Program Files/Git/git-bash.exe, D:/Monero/wallets/%Wallet%1
    Sleep, 1000
    Send, comm  -1 -3 --nocheck-order  txs-old.csv txs.csv > diff.csv{Enter} 
    Sleep, 1000
    Send, exit{Enter}
    FileDelete, D:\Monero\wallets\Common\diff.csv
    FileMove, D:\Monero\wallets\%Wallet%1\diff.csv, D:\Monero\wallets\Common\diff.csv
    FileDelete, D:\Monero\wallets\%Wallet%1\txs-old-old-old.csv
    FileMove, D:\Monero\wallets\%Wallet%1\txs-old-old.csv, D:\Monero\wallets\%Wallet%1\txs-old-old-old.csv
    FileMove, D:\Monero\wallets\%Wallet%1\txs-old.csv, D:\Monero\wallets\%Wallet%1\txs-old-old.csv
    FileMove, D:\Monero\wallets\%Wallet%1\txs.csv, D:\Monero\wallets\%Wallet%1\txs-old.csv
    Sleep, 500
    Run, D:\Excel\Book1.xlsm
    Sleep, 500
    WinActivate, Microsoft Excel - Book1.xlsm  [Read-Only]
    Sleep, 333
    Send, !{F8}
    Sleep, 500
    Send, {Enter}
    Sleep, 500
    XL := ComObjActive("Excel.Application") 
    If (!IsObject(XL))
        XL := ComObjCreate("Excel.Application")
    XL.Range("B1").Select
    Inout := XL.ActiveCell.Text
    While Inout !=
    {
        Inout := XL.ActiveCell.Text
        IfInString, Inout, in
        {
            XL.ActiveCell.Offset(0, 2).Select
            Time := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 1).Select
            ReportAmount := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 1).Select
            Balance := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 10).Select
            Name := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 1).Select
            Don := XL.ActiveCell.Value
            Amount := Round(Don, 2)
            XL.ActiveCell.Offset(0, 2).Select
            ReportName := XL.ActiveCell.Text
            /*
            MsgBox, 0, , 
            (LTrim
            RN %ReportName%
            Time %Time%
            RA%ReportAmount%
            Bal  %Balance%
            Name %Name%
            Amount %Amount%
            )
            */
            FileAppend, %ReportName%   %Time%    %ReportAmount%    %Balance%{Enter}, D:\Monero\wallets\%Wallet%1\Balance.txt
            WinActivate, AvengersAustralia  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
            Sleep, 333
            Sleep, 300
            WinMenuSelectItem, AvengersAustralia  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant), , 6&, 2&
            Sleep, 2000
            Send, %Name%{Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, Donation{Tab}
            Sleep, 1000
            Send, %Time%{Tab}
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, %Amount%
            Sleep, 1000
            Send, {Tab}
            Sleep, 1000
            Send, 
            /*
            Send, {Enter}
            */
            Send, !a
            XL.ActiveCell.Offset(1, -17).Select
            Sleep, 1000
            Continue
        }
        Else IfInString, Inout, out
        {
            XL.ActiveCell.Offset(0, 2).Select
            Date := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 1).Select
            ReportAmount := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 1).Select
            Balance := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 11).Select
            Don := XL.ActiveCell.Value
            Amount := Round(Don, 2)
            FileAppend, Test Payment    %Date%    -%ReportAmount%    %Balance% {Enter}, D:\Monero\wallets\%Wallet%1\Balance.txt
            WinActivate, AvengersAustralia  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
            Sleep, 333
            Sleep, 300
            Send, ^w
            Sleep, 2000
            Send, Funding{Enter}
            Sleep, 500
            Send, {Tab}
            Sleep, 500
            Send, {Tab}
            Sleep, 500
            Sleep, 500
            Send, {Tab}
            Sleep, 500
            Send, %Amount%{Tab}
            Sleep, 500
            Send, %Date%
            Send, {Tab}
            Send, !m
            Sleep, 500
            Send, Test
            Sleep, 500
            Send, {Tab}
            Sleep, 500
            Send, {Tab}
            Sleep, 500
            Send, 1{Tab}
            Sleep, 500
            Send, %Amount%{Tab}
            Sleep, 500
            Send, !a
            Sleep, 500
            XL := ComObjActive("Excel.Application")
            XL.ActiveCell.Offset(1, -15).Select
            Continue
        }
    }
    Else
    {
        WinActivate, Microsoft Excel
        Sleep, 333
        WinClose, Microsoft Excel
        Sleep, 333
        Sleep, 500
        SendRaw, n
        Sleep, 1000
        MsgBox, 262144, , Turn off Excel
        Goto, Outer
    }
    Until, Wallet := "z"
}
WinWait, QuickBooks Desktop Login
Sleep, 333
Return

