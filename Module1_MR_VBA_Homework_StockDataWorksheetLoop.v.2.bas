Attribute VB_Name = "Module1"
 Sub WorksheetsLoop()

          
            Dim CurrentWs As Worksheet
            Dim Summary_Table_Header As Boolean
            Summary_Table_Header = True
            Active_Worksheet = True
            
            ' Loop through all of the worksheets
            For Each CurrentWs In Worksheets
            
                ' Set initial variable for the ticker name and toal per ticket name
                Dim Ticker_Name As String
                Ticker_Name = " "
                Dim Total_Ticker_Volume As Double
                Total_Ticker_Volume = 0
                
                ' Set additional variables
                Dim Open_Price As Double
                Open_Price = 0
                Dim Close_Price As Double
                Close_Price = 0
                Dim YearlyChange_Price As Double
                YearlyChange_Price = 0
                Dim YearlyChange_Percent As Double
                YearlyChange_Percent = 0
                Dim Max_Ticker_Name As String
                Max_Ticker_Name = " "
                Dim Min_Ticker_Name As String
                Min_Ticker_Name = " "
                Dim Max_Percent As Double
                Max_Percent = 0
                Dim Min_Percent As Double
                Min_Percent = 0
                Dim Max_Volume_Ticker As String
                Max_Volume_Ticker = " "
                Dim Max_Volume As Double
                Max_Volume = 0
                Dim Summary_Table_Row As Long
                Summary_Table_Row = 2
                
                ' Set initial row count for the current worksheet
                Dim Lastrow As Long
                Dim i As Long
                
                Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

                ' insert headers and titles
                 Range("I1").Value = "Ticker"
                 Range("J1").Value = "Yearly Change"
                 Range("K1").Value = "Percent Change"
                 Range("L1").Value = "Total Stock Volume"

                 Range("O2").Value = "Greatest % Increase"
                 Range("O3").Value = "Greatest % Decrease"
                 Range("O4").Value = "Greatest Total Volume"

                 Range("P1").Value = "Ticker"
                 Range("Q1").Value = "Value"

                
                ' Set initial value of open price for the first ticker and then start a loop
                Open_Price = CurrentWs.Cells(2, 3).Value
                
                ' Loop from the beginning to lastrow of current worksheet
                For i = 2 To Lastrow
                
                    ' if still within the same ticker name and if not add it to the summary table
                    
                    If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                    
                        Ticker_Name = CurrentWs.Cells(i, 1).Value
                        Close_Price = CurrentWs.Cells(i, 6).Value
                        YearlyChange_Price = Close_Price - Open_Price
    
                        If Open_Price <> 0 Then
                            YearlyChange_Percent = (YearlyChange_Price / Open_Price) * 100
                        Else
                            
                        End If
                        
                        Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                      
                        
                        ' Print the Ticker Name in the Summary Table
                        CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                        CurrentWs.Range("J" & Summary_Table_Row).Value = YearlyChange_Price
                        ' Add colors to Yearly Change with Green and Red colors
                        If (YearlyChange_Price > 0) Then
                            'Fill column with Green color
                            CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        ElseIf (YearlyChange_Price <= 0) Then
                            'Fill column with Red color
                            CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                        
                         ' Print the Ticker Name in the Summary Table
                        CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(YearlyChange_Percent) & "%")
                        CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                        
                        ' Add 1 to the summary table row count
                        Summary_Table_Row = Summary_Table_Row + 1
                        ' Reset YearlyChange_Pice and Yearly Change_Percent
                        YearlyChange_Price = 0
                        Close_Price = 0
                        ' Look for next Ticker's Open_Price
                        Open_Price = CurrentWs.Cells(i + 1, 3).Value

                        If (YearlyChange_Percent > Max_Percent) Then
                            Max_Percent = YearlyChange_Percent
                            Max_Ticker_Name = Ticker_Name
                        ElseIf (YearlyChange_Percent < Min_Percent) Then
                            Min_Percent = YearlyChange_Percent
                            Min_Ticker_Name = Ticker_Name
                        End If
                               
                        If (Total_Ticker_Volume > Max_Volume) Then
                            Max_Volume = Total_Ticker_Volume
                            Max_Volume_Ticker = Ticker_Name
                        End If
                        
                        ' Reset Yearly Change Percent and Yearly Change Volume
                        YearlyChange_Percent = 0
                        Total_Ticker_Volume = 0

                    Else
                        ' Increase the Total Ticker Volume
                        Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                    End If
              
                Next i

                    If Not Active_Worksheet Then
                    
                        CurrentWs.Range("Q2").Value = (CStr(Max_Percent) & "%")
                        CurrentWs.Range("Q3").Value = (CStr(Min_Percent) & "%")
                        CurrentWs.Range("P2").Value = Max_Ticker_Name
                        CurrentWs.Range("P3").Value = Min_Ticker_Name
                        CurrentWs.Range("Q4").Value = Max_Volume
                        CurrentWs.Range("P4").Value = Max_Volume_Ticker
                        
                    Else
                        Active_Worksheet = False
                    End If
                
             Next CurrentWs
    End Sub

