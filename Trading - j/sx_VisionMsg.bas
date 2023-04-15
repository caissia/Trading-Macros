Attribute VB_Name = "sx_VisionMsg"
Option Explicit

Sub All_Vision()
Attribute All_Vision.VB_ProcData.VB_Invoke_Func = "U\n14"
'messages of focus, tenacity, perseverance

    Dim LastMessage As String
    Dim NextMessage As String
    Dim PerMessage As String
    Dim StMessage As String
    Dim temp As Variant
    Dim m9(9) As String
    Dim n(8) As Variant
    Dim m(8) As String
    Dim i As Long
    Dim r As Long

    'permanent address of next cell address
    PerMessage = "$U$35"

    'start cell of randomized order of message numbers
    StMessage = "$V$34"

    'contains the cell address of next message
    NextMessage = Sheets("Range").Range(PerMessage)

    'contains the cell address of last message given
    LastMessage = "$V$42"

    'populate messages
    GoSub Messages

    'give current message
    temp = Sheets("range").Range(NextMessage)
    If temp <> 9 Then
        MsgBox m(temp), 0, "Ultimate Space: " & temp
    Else
        For i = 1 To 8
            MsgBox m9(i), 0, "Ultimate Space: 9." & i
        Next i
    End If

    'record the next message number
    Sheets("Range").Range(PerMessage) = Sheets("Range").Range(NextMessage).Offset(1).Address

    'if last message was given then create/store randomized order of messages again
    If LastMessage = NextMessage Then

        'create array of 8 items
        For i = 0 To 8
            n(i) = i + 1
        Next i

        'randomize that array
        Randomize
        For i = LBound(n) To UBound(n)
            r = CLng(((UBound(n) - i) * Rnd) + i)
            If i <> r Then
                temp = n(i)
                n(i) = n(r)
                n(r) = temp
            End If
        Next i

        'paste the newly randomized order of messages
        For i = 0 To 8
            Sheets("Range").Range(StMessage).Offset(i) = n(i)
        Next i

        'reset the last message address to the first address
        Sheets("Range").Range(PerMessage) = StMessage

    End If


Exit Sub


Messages:

    m(1) = "If a man does not keep pace with his companions," & Chr(10) _
           & "  perhaps it is because he hears a different drummer." & Chr(10) & Chr(10) _
           & "Let him step to the music which he hears," & Chr(10) _
           & "  however measured or far away."
           '& "  - Thoreau"

    m(2) = "***Trade with the Trend***" & Chr(10) _
           & "     ---on every trade---" & Chr(10) _
           & "***Let your trades work***" & Chr(10) _
           & "       -Daily… to Daily -" & Chr(10) _
           & "         - 4hr… to 4hr -" & Chr(10) _
           & "         - 1hr… to 1hr -"

    m(3) = "*** No trading impulsively ***" & Chr(10) _
           & "         ---on any trade---" & Chr(10) _
           & "  *** Let your trades work ***" & Chr(10) _
           & "         - Daily… to Daily -" & Chr(10) _
           & "            - 4hr… to 4hr -" & Chr(10) _
           & "            - 1hr… to 1hr -"

    m(4) = "This is the trading plan..." & Chr(10) _
           & "  The key here for me is complete and accurate" & Chr(10) _
           & "  daily records so that I can assess my trading," & Chr(10) _
           & "  review my progress on a month-by-month basis and" & Chr(10) _
           & "  make changes in strategies based on my performance."
           '& "  - Carter"

    m(5) = "The fun in trading comes from the" & Chr(10) _
           & "  thrill of the hunt, the anticipation of the kill." & Chr(10) & Chr(10) _
           & "All the research, all the work" & Chr(10) _
           & "  culminates into a single moment in time" & Chr(10) _
           & "  when a trader makes a decision to pull the trigger" & Chr(10) _
           & "  and is shortly thereafter presented with the results."
           '& "  - Carter" 'http://www.tradethemarkets.com/public/Online_Stock_Trading_Plan.cfm

    m(6) = "The Best Habit of All" & Chr(10) & Chr(10) _
           & "  Believe in yourself." & Chr(10) _
           & "  Confidence is the most powerful trading tool." & Chr(10) _
           & "  Experiments have shown that people who believe they have an" & Chr(10) _
           & "  edge often perform better, even when that belief was unfounded." & Chr(10) _
           & "  In many cases, thinking that we are limited is itself a limiting factor." & Chr(10) & Chr(10) _
           & "  There is accumulating evidence that suggests that our thoughts are" & Chr(10) _
           & "  often capable of extending our cognitive and physical limits." & Chr(10) _
           & "  If you are going to make a trade trust yourself." & Chr(10) & Chr(10) _
           & "  If you are watching markets closely enough" & Chr(10) _
           & "  to be reading this on a summer afternoon" & Chr(10) _
           & "  you’ve probably put in the work and that" & Chr(10) _
           & "  puts you ahead of 90% of traders." & Chr(10) _
           & "  Clear your head and go for it."
           '& "  - Adam Button" 'August 15th, 2013 18:59:10 GMT

    m(7) = "~~~~~~~~~~~~~~~~~~~~~" & Chr(10) _
           & "         Only those who dare" & Chr(10) _
           & "              to fail miserably" & Chr(10) _
           & "           can achieve greatly." & Chr(10) _
           & "~~~~~~~~~~~~~~~~~~~~~"
           '& " - J. F. Kennedy"

    m(8) = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & Chr(10) _
           & "         The future rewards those who press on." & Chr(10) _
           & "        I don't have time to feel sorry for myself." & Chr(10) _
           & "                I don't have time to complain." & Chr(10) _
           & "                      I'm going to press on." & Chr(10) _
           & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
           '& " - B. H. Obama"

    m9(1) = "Pull the Trigger:" & Chr(10) & Chr(10) _
           & "  Taking the trade for each valid setup is" & Chr(10) _
           & "  a small step towards success and wealth." & Chr(10) _
           & "  Taking the trade on each valid setup carries" & Chr(10) _
           & "  risk but it is also the only way to success."

    m9(2) = "Every Moment Is Unique:" & Chr(10) & Chr(10) _
           & "  The trade either works or it doesn't."

    m9(3) = "Anything Can Happen:" & Chr(10) & Chr(10) _
           & "  Develop a resolute, unshakeable belief in uncertainty." & Chr(10) _
           & "  The market has no responsibility to give us anything" & Chr(10) _
           & "  or do anything that would benefit us."

    m9(4) = "Losses are part of Trading:" & Chr(10) & Chr(10) _
           & "  Losing and being wrong are inevtiable" & Chr(10) _
           & "  realities of trading since anything can happen." & Chr(10) _
           & "  Taking small losses is part of a succesful trader's job."

    m9(5) = "Accept Risk:" & Chr(10) & Chr(10) _
           & "  Fully acknowledge the risks inherent in trading" & Chr(10) _
           & "  and accept complete responsibility for each" & Chr(10) _
           & "  trade (not the market). When a loss occurs, do" & Chr(10) _
           & "  not suffer emotional discomfort or fear."

    m9(6) = "Monitor Emotions:" & Chr(10) & Chr(10) _
           & "  Learn how to monitor and control the negative effects " & Chr(10) _
           & "  of euphoria and the potential for self-sabotage."

    m9(7) = "Abandon Search for Holy Grail:" & Chr(10) & Chr(10) _
           & "  Attitude produces better overall results" & Chr(10) _
           & "  than analysis or technique."

    m9(8) = "Rigid Rules, Flexible Expectations:" & Chr(10) & Chr(10) _
           & "  Adopt rigidity in your trading rules" & Chr(10) _
           & "  and flexibility in your expectations."

Return


End Sub
