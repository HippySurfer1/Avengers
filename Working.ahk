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
Sleep, 40000
WinWait, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Accountant 18.0, , 15000
Sleep, 333
WinMenuSelectItem, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Accountant 18.0 ahk_class MauiFrame ahk_exe QBW32.EXE, , File, Toggle to Another Edition
Sleep, 50
Send, {Enter}
Sleep, 2000
WinActivate, Select QuickBooks Desktop Industry-Specific Edition
Sleep, 333
Send, {Down}{Down}{Down}
Send, {Space}
Send, {Enter}
Send, {Enter}
WinClose, Automatic Backup, , 5
Sleep, 333
WinWait, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
Sleep, 333
/*
Send, {Enter}
*/
Run, C:/Windows/SysWOW64/WindowsPowerShell/v1.0/powershell.exe, C:/Monero
Sleep, 500
Send, ./monerod.exe --data-dir E:/BitMonero {Enter}
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
    Sleep, 300
    Send, refresh {Enter}
    MsgBox, 262144, , Please click OK when the wallet is synced
    Send, export_transfers {Enter}
    Sleep, 300
    Send, exit{Enter}
    Sleep, 300
    Send, exit{Enter}
    FileMove, C:/Monero/output0.csv, D:/Monero/wallets/%Wallet%1/txs.csv
    Run, C:/Program Files/Git/git-bash.exe, D:/Monero/wallets/%Wallet%1
    Sleep, 3000
    Send, comm  -1 -3 --nocheck-order  txs-old.csv txs.csv > diff.csv{Enter} 
    Sleep, 3000
    Send, exit{Enter}
    FileDelete, D:\Monero\wallets\Common\diff.csv
    FileMove, D:\Monero\wallets\%Wallet%1\diff.csv, D:\Monero\wallets\Common\diff.csv
    FileDelete, D:\Monero\wallets\%Wallet%1\txs-old-old-old.csv
    FileMove, D:\Monero\wallets\%Wallet%1\txs-old-old.csv, D:\Monero\wallets\%Wallet%1\txs-old-old-old.csv
    FileMove, D:\Monero\wallets\%Wallet%1\txs-old.csv, D:\Monero\wallets\%Wallet%1\txs-old-old.csv
    FileMove, D:\Monero\wallets\%Wallet%1\txs.csv, D:\Monero\wallets\%Wallet%1\txs-old.csv
    Sleep, 300
    Run, D:\Excel\Book1.xlsm
    Sleep, 300
    WinActivate, Microsoft Excel - Book1.xlsm  [Read-Only]
    Sleep, 333
    Send, !{F8}
    Sleep, 300
    Send, {Enter}
    Sleep, 300
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
            XL.ActiveCell.Offset(0, 13).Select
            Don := XL.ActiveCell.Value
            Amount := Round(Don, 2)
            XL.ActiveCell.Offset(0, -1).Select
            Name := XL.ActiveCell.Text
            WinActivate, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
            Sleep, 333
            WinMenuSelectItem, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant), , 6&, 2&
            Sleep, 5000
            Send, %Name%{Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 200
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, Donation{Tab}
            Sleep, 300
            Send, %Time%{Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, %Amount%
            Send, {Tab}
            Sleep, 300
            Send, 
            /*
            Send, {Enter}
            */
            Send, !a
            XL.ActiveCell.Offset(1, -14).Select
            Sleep, 3000
            Continue
        }
        Else IfInString, Inout, out
        {
            XL.ActiveCell.Offset(0, 2).Select
            Date := XL.ActiveCell.Text
            XL.ActiveCell.Offset(0, 13).Select
            Don := XL.ActiveCell.Value
            Amount := Round(Don, 2)
            WinActivate, AvengersAu  - Intuit QuickBooks Enterprise Solutions: Nonprofit 18.0 (via Accountant)
            Sleep, 333
            Send, ^w
            Sleep, 1000
            Send, Fund{Enter}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, %Amount%{Tab}
            Sleep, 300
            Send, %Date%
            Send, {Tab}
            Send, !m
            Sleep, 300
            Send, Test
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, {Tab}
            Sleep, 300
            Send, 1{Tab}
            Sleep, 300
            Send, %Amount%{Tab}
            Sleep, 300
            Send, !a
            Sleep, 300
            XL := ComObjActive("Excel.Application")
            XL.ActiveCell.Offset(1, -15).Select
            Continue
        }
        WinClose, Microsoft Excel - Book1.xlsm
        Sleep, 333
        Sleep, 300
        SendRaw, n
        Goto, Outer
    }
    Until, Wallet := "x"
}
Return

F4::
Macro2:
FileDelete, D:\Monero\wallets\Avengers1\txs-old.csv
FileDelete, D:\Monero\wallets\Larry1\txs-old.csv
FileDelete, D:\Monero\wallets\Curley1\txs-old.csv
FileDelete, D:\Monero\wallets\Moe1\txs-old.csv
FileCopy, D:\Monero\wallets\txs-old.csv, D:\Monero\wallets\Avengers1\txs-old.csv
FileCopy, D:\Monero\wallets\txs-old.csv, D:\Monero\wallets\Curley1\txs-old.csv
FileCopy, D:\Monero\wallets\txs-old.csv, D:\Monero\wallets\Larry1\txs-old.csv
FileCopy, D:\Monero\wallets\txs-old.csv, D:\Monero\wallets\Moe1\txs-old.csv
Return

