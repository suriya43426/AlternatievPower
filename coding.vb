Sub sb_new()
   
    Dim conn As New ADODB.Connection
    Dim iRowNo, iRowNo1 As Integer
    Dim UpdateSql, UpdateSql2, sDate, sDate2, sCompany, sBlock, sGen, sBoiler, sProduct, sUnit, sCheck, sType As String
    Dim sFCheck, sFProduct, sFBlock, sFUnit, sFBoiler, sReceiver, sRemark, sNo, sCheck2, DATETIME As String
    Dim sAct_qty0, sAct_qty, sPlan_qty, sPlan_qty0 As Double
    Dim sAct_qty01, sAct_qty1, sPlan_qty1, sPlan_qty01 As Double
    Dim sAct_qtyf, sAct_qty0f, sPlan_qtyf, sPlan_qty0f As Double
  
    With Sheets("Input2")
            
        'Open a connection to SQL Server
        conn.Open "Provider=SQLOLEDB;Data Source=10.25.69.142;Initial Catalog=findb_prd;User ID=finuser;Password=f!nP@ssW0rd;"
        'conn.Open "Provider=SQLOLEDB;Data Source=MPGDC-CDB;Initial Catalog=findb_prd;User ID=finuser;Password=f!nP@ssW0rd;"
        'Skip the header row
        iRowNo = 8

        'Loop until empty cell in CustomerId
        Do Until .Cells(iRowNo, 1) = ""
            sDate = .Cells(4, 3)
            sDate2 = .Cells(4, 4)
            sCompany = .Cells(4, 2)
            sBlock = .Cells(iRowNo, 4)
            sGen = .Cells(iRowNo, 3)
            sBoiler = .Cells(iRowNo, 6)
            sProduct = .Cells(iRowNo, 5)
            sUnit = .Cells(iRowNo, 13)
            sType = .Cells(iRowNo, 7)
            sAct_qty = .Cells(iRowNo, 15)
            sPlan_qty = 0
            sAct_qtyf = .Cells(iRowNo, 16)
            sPlan_qtyf = 0
            sCustomer = .Cells(iRowNo, 8)
            sCheck = .Cells(iRowNo, 2)
            sRemark = .Cells(iRowNo, 12)
            sRemark2 = .Cells(iRowNo, 12)
            sReceiver = .Cells(iRowNo, 9)
            sNo = .Cells(iRowNo, 10)
            sCheck2 = .Cells(iRowNo, 20)
            DATETIME = Format(Now, "yyyy-mm-dd hh:mm:ss")
            'insert test
           
            
        'check Act qty and Plan qty set1
        If sAct_qty = "" Or sAct_qty = 0 Then
            sAct_qty0 = 0
        Else
            sAct_qty0 = sAct_qty
        End If

        If sPlan_qty = "" Or sPlan_qty = 0 Then
            sPlan_qty0 = 0
        Else
            sPlan_qty0 = sPlan_qty
        End If
        If sPlan_qtyf = "" Or sPlan_qtyf = 0 Then
            sPlan_qty0f = 0
        Else
            sPlan_qty0f = sPlan_qtyf
        End If
         
        If sAct_qtyf = "" Or sAct_qtyf = 0 Then
            sAct_qty0f = 0
        Else
            sAct_qty0f = sAct_qtyf
        End If
    
       
            
'''''''''''''''''''''''''''''''''''''SQL Insert&Update Database table'''''''''''''''''''''''''''''''''''''
    'Table Production
     If sCheck = "P" And sCheck2 = "T" Then
        UpdateSql = "UPDATE dbo.bio_production " & _
        "SET date= '" & sDate & "'," & _
        "company= '" & sCompany & "'," & _
        "block= '" & sBlock & "'," & _
        "generator= '" & sGen & "'," & _
        "boiler= '" & sBoiler & "'," & _
        "product= '" & sProduct & "'," & _
        "unit= '" & sUnit & "'," & _
        "type= '" & sType & "'," & _
        "act_qty= " & sAct_qty0 & "," & _
        "plan_qty= " & sPlan_qty0 & "," & _
        "DATETIME= '" & DATETIME & "'" & _
        "WHERE date= '" & sDate & "' and " & _
        "company= '" & sCompany & "' and " & _
        "block= '" & sBlock & "' and " & _
        "generator= '" & sGen & "' and " & _
        "boiler= '" & sBoiler & "' and " & _
       "product= '" & sProduct & "' and " & _
        "unit= '" & sUnit & "' and " & _
        "type= '" & sType & "'"



        InsertSql = "INSERT INTO dbo.bio_production VALUES ('" & _
        sDate & "'," & _
        "'" & sCompany & "'," & _
        "'" & sBlock & "'," & _
        "'" & sGen & "'," & _
        "'" & sBoiler & "'," & _
        "'" & sProduct & "'," & _
        "'" & sUnit & "'," & _
        "'" & sType & "'," & _
        "" & sAct_qty0 & "," & _
        "" & sPlan_qty0 & "," & _
        "'" & DATETIME & "'" & _
        ")"

    'Table Disribution
    ElseIf sCheck = "D" And sCheck2 = "T" Then
        UpdateSql = "UPDATE dbo.bio_dist " & _
        "SET date= '" & sDate & "'," & _
        "company= '" & sCompany & "'," & _
        "block= '" & sBlock & "'," & _
        "product= '" & sProduct & "'," & _
        "unit= '" & sUnit & "'," & _
        "type= '" & sType & "'," & _
        "customer= '" & sCustomer & "'," & _
        "act_qty= " & sAct_qty0 & "," & _
        "plan_qty= " & sPlan_qty0 & "," & _
        "DATETIME= '" & DATETIME & "'" & _
        "WHERE date= '" & sDate & "' and " & _
        "company= '" & sCompany & "' and " & _
        "block= '" & sBlock & "' and " & _
        "product= '" & sProduct & "' and " & _
        "unit= '" & sUnit & "' and " & _
        "type= '" & sType & "' and " & _
        "customer= '" & sCustomer & "'"



        InsertSql = "INSERT INTO dbo.bio_dist VALUES ('" & _
        sDate & "'," & _
        "'" & sCompany & "'," & _
        "'" & sProduct & "'," & _
        "'" & sType & "'," & _
        "'" & sBlock & "'," & _
        "'" & sUnit & "'," & _
        "'" & sCustomer & "'," & _
        "" & sAct_qty0 & "," & _
        "" & sPlan_qty0 & "," & _
        "'" & DATETIME & "'" & _
        ")"
        
        
        'Table Fuel Used
                ElseIf sCheck = "FU" And sCheck2 = "T" Then
                    UpdateSql = "UPDATE dbo.bio_fuel_used " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "product= '" & sProduct & "'," & _
                    "block= '" & sBlock & "'," & _
                    "boiler= '" & sBoiler & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "plan_qty= " & sPlan_qty0 & "," & _
                    "act_qty= " & sAct_qty0 & "," & _
                    "plan_factor_qty= " & sPlan_qty0f & "," & _
                    "act_factor_qty= " & sAct_qty0f & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "block= '" & sBlock & "' and " & _
                    "product= '" & sProduct & "' and " & _
                    "unit= '" & sUnit & "' and " & _
                    "boiler= '" & sBoiler & "' "
            
            

                    InsertSql = "INSERT INTO dbo.bio_fuel_used VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sProduct & "'," & _
                    "'" & sBlock & "'," & _
                    "'" & sBoiler & "'," & _
                    "'" & sUnit & "'," & _
                    "" & sPlan_qty0 & "," & _
                    "" & sAct_qty0 & "," & _
                    "" & sPlan_qty0f & "," & _
                    "" & sAct_qty0f & "," & _
                    "'" & DATETIME & "'" & _
                    ")"

            
                'Table Fuel Receiver
                ElseIf sCheck = "FR" And sCheck2 = "T" Then
                    UpdateSql = "UPDATE dbo.bio_fuel_rec " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "Receiver= '" & sReceiver & "'," & _
                    "product= '" & sProduct & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "plan_qty= " & sPlan_qty0 & "," & _
                    "act_qty= " & sAct_qty0 & "," & _
                    "plan_factor_qty= " & sPlan_qty0f & "," & _
                    "act_factor_qty= " & sAct_qty0f & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "Receiver= '" & sReceiver & "' and " & _
                    "product= '" & sProduct & "' and " & _
                    "unit= '" & sUnit & "' "




                    InsertSql = "INSERT INTO dbo.bio_fuel_rec VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                   "'" & sProduct & "'," & _
                    "'" & sReceiver & "'," & _
                    "'" & sUnit & "'," & _
                    "" & sPlan_qty0 & "," & _
                    "" & sAct_qty0 & "," & _
                    "" & sPlan_qty0f & "," & _
                    "" & sAct_qty0f & "," & _
                    "'" & DATETIME & "'" & _
                    ")"

                'Table Fuel Stock
               ElseIf sCheck = "FS" And sCheck2 = "T" Then
                    UpdateSql = "UPDATE dbo.bio_fuel_stock " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "product= '" & sProduct & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "plan_qty= " & sPlan_qty0 & "," & _
                    "act_qty= " & sAct_qty0 & "," & _
                    "plan_factor_qty= " & sPlan_qty0f & "," & _
                    "act_factor_qty= " & sAct_qty0f & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "product= '" & sProduct & "' and " & _
                    "unit= '" & sUnit & "' "



                    InsertSql = "INSERT INTO dbo.bio_fuel_stock VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sProduct & "'," & _
                    "'" & sUnit & "'," & _
                    "" & sPlan_qty0 & "," & _
                    "" & sAct_qty0 & "," & _
                    "" & sPlan_qty0f & "," & _
                    "" & sAct_qty0f & "," & _
                    "'" & DATETIME & "'" & _
                    ")"

                'Table Fuel Water
                ElseIf sCheck = "W" And sCheck2 = "T" Then
                    UpdateSql = "UPDATE dbo.bio_fuel_wateru " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "block= '" & sBlock & "'," & _
                    "product= '" & sProduct & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "type= '" & sType & "'," & _
                    "no= '" & sNo & "'," & _
                    "plan_qty= " & sPlan_qty0 & "," & _
                    "act_qty= " & sAct_qty0 & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "block= '" & sBlock & "' and " & _
                    "product= '" & sProduct & "' and " & _
                    "unit= '" & sUnit & "' and " & _
                    "type= '" & sType & "' and " & _
                    "no= '" & sNo & "' "


                   InsertSql = "INSERT INTO dbo.bio_fuel_wateru VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sBlock & "'," & _
                    "'" & sProduct & "'," & _
                    "'" & sUnit & "'," & _
                    "'" & sType & "'," & _
                    "'" & sNo & "'," & _
                    "" & sPlan_qty0 & "," & _
                    "" & sAct_qty0 & "," & _
                    "'" & DATETIME & "'" & _
                    ")"

                'Table Fuel Quality
                ElseIf sCheck = "Q" And sCheck2 = "T" Then
                   UpdateSql = "UPDATE dbo.bio_fuel_quality " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "product= '" & sProduct & "'," & _
                    "block= '" & sBlock & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "type= '" & sType & "'," & _
                    "act_qty= " & sAct_qty0 & "," & _
                    "plan_qty= " & sPlan_qty0 & "," & _
                    "act_used_qty= " & sAct_qty0f & "," & _
                    "plan_used_qty= " & sPlan_qty0f & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "product= '" & sProduct & "' and " & _
                    "block= '" & sBlock & "' and " & _
                    "unit= '" & sUnit & "' and " & _
                    "type= '" & sType & "' "




                    InsertSql = "INSERT INTO dbo.bio_fuel_quality VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sProduct & "'," & _
                    "'" & sBlock & "'," & _
                    "'" & sUnit & "'," & _
                    "'" & sType & "'," & _
                    "" & sAct_qty0 & "," & _
                    "" & sPlan_qty0 & "," & _
                    "" & sAct_qty0f & "," & _
                    "" & sPlan_qty0f & "," & _
                    "'" & DATETIME & "'" & _
                    ")"
                    
                'Table Efficiency
                ElseIf sCheck = "E" And sCheck2 = "T" Then
                    UpdateSql = "UPDATE dbo.bio_efficiency " & _
                   "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "block= '" & sBlock & "'," & _
                    "boiler= '" & sBoiler & "'," & _
                    "product= '" & sProduct & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "act_qty= " & sAct_qty0 & "," & _
                    "plan_qty= " & sPlan_qty0 & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "block= '" & sBlock & "' and " & _
                    "boiler= '" & sBoiler & "' and " & _
                    "product= '" & sProduct & "' and " & _
                    "unit= '" & sUnit & "'"



                    InsertSql = "INSERT INTO dbo.bio_efficiency VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sBlock & "'," & _
                    "'" & sBoiler & "'," & _
                    "'" & sProduct & "'," & _
                    "'" & sUnit & "'," & _
                    "" & sAct_qty0 & "," & _
                    "" & sPlan_qty0 & "," & _
                    "'" & DATETIME & "'" & _
                    ")"
                    
                    'Table Cane
                ElseIf sCheck = "C" And sCheck2 = "T" Then
                    UpdateSql = "UPDATE dbo.bio_cane " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "product= '" & sProduct & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "act_qty= " & sAct_qty0 & "," & _
                    "plan_qty= " & sPlan_qty0 & "," & _
                    "act_cane_qty= " & sAct_qty0f & "," & _
                    "plan_cane_qty= " & sPlan_qty0f & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "product= '" & sProduct & "' and " & _
                    "unit= '" & sUnit & "'"



                    InsertSql = "INSERT INTO dbo.bio_cane VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sProduct & "'," & _
                    "'" & sUnit & "'," & _
                    "" & sAct_qty0 & "," & _
                    "" & sPlan_qty0 & "," & _
                    "" & sAct_qty0f & "," & _
                    "" & sPlan_qty0f & "," & _
                    "'" & DATETIME & "'" & _
                    ")"
        'Table Remark
                ElseIf sCheck = "R" And sCheck2 = "T" Then
                    
                    UpdateSql = "UPDATE dbo.bio_remark " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "remark= '" & sRemark & "'," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "'"
            
            
            
                    InsertSql = "INSERT INTO dbo.bio_remark VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sRemark & "'," & _
                    "'" & DATETIME & "'" & _
                    ")"
                    
                    
                            'Table Remark
                ElseIf sCheck = "R2" And sCheck2 = "T" Then
                    
                    UpdateSql = "UPDATE dbo.bio_remark_cane_leaves " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "remark= '" & sRemark2 & "'," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "'"
            
            
            
                    InsertSql = "INSERT INTO dbo.bio_remark_cane_leaves VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sRemark2 & "'," & _
                    "'" & DATETIME & "'" & _
                    ")"
                    
                    'Table Fuel Carry Forward
                           ElseIf sCheck = "CR" And sCheck2 = "T" Then
                                UpdateSql = "UPDATE dbo.bio_fuel_carry_fwd " & _
                                "SET date= '" & sDate2 & "'," & _
                                "company= '" & sCompany & "'," & _
                                "product= '" & sProduct & "'," & _
                                "unit= '" & sUnit & "'," & _
                                "type= '" & sType & "'," & _
                                "act_qty= " & sAct_qty0 & "" & _
                                "WHERE date= '" & sDate2 & "' and " & _
                                "company= '" & sCompany & "' and " & _
                                "product= '" & sProduct & "' and " & _
                                "unit= '" & sUnit & "' and " & _
                                "type= '" & sType & "' "
            
            
            
                                InsertSql = "INSERT INTO dbo.bio_fuel_carry_fwd VALUES ('" & _
                                sDate2 & "'," & _
                                "'" & sCompany & "'," & _
                                "'" & sProduct & "'," & _
                                "'" & sUnit & "'," & _
                                "'" & sType & "'," & _
                                "" & sAct_qty0 & "" & _
                                ")"
                    
     
      'Table downtime
            ElseIf sCheck = "DW" And sCheck2 = "T" Then
            
                    UpdateSql = "UPDATE dbo.bio_downtime " & _
                    "SET date= '" & sDate & "'," & _
                    "company= '" & sCompany & "'," & _
                    "item= '" & sRemark & "'," & _
                    "unit= '" & sUnit & "'," & _
                    "actual= " & sAct_qty & "," & _
                    "DATETIME= '" & DATETIME & "'" & _
                    "WHERE date= '" & sDate & "' and " & _
                    "company= '" & sCompany & "' and " & _
                    "item= '" & sRemark & "' and " & _
                    "unit= '" & sUnit & "'"
            
            
            
                    InsertSql = "INSERT INTO dbo.bio_downtime VALUES ('" & _
                    sDate & "'," & _
                    "'" & sCompany & "'," & _
                    "'" & sRemark & "'," & _
                    "'" & sUnit & "'," & _
                    "" & sAct_qty & "," & _
                    "'" & DATETIME & "'" & _
                    ")"
    
    End If


              
If UpdateSql <> "" Then
    conn.Execute UpdateSql, RecordsAffected
    
    End If
    
If InsertSql <> "" Then
    If RecordsAffected = 0 Then
        conn.Execute InsertSql
      
    End If
    
    End If
    
    
        
 
            iRowNo = iRowNo + 1
            'Exit Do
        Loop
    
        
    


              

            
            
        MsgBox "Completed."
            
        conn.Close
        Set conn = Nothing
             
    End With
End Sub

