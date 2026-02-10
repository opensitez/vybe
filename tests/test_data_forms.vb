' Test Data Forms: BindingSource, DataGridView DataSource, DataBindings
' Tests the complete data form binding pipeline

Imports System
Imports System.Data
Imports System.Windows.Forms

Module TestDataForms

    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, description As String)
        If condition Then
            Console.WriteLine("PASS: " & description)
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: " & description)
            Console.WriteLine("FAILURE")
            failed = failed + 1
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== Data Forms Tests ===")
        Console.WriteLine()

        ' ---- Test 1: Create BindingSource ----
        Console.WriteLine("--- BindingSource Creation ---")
        Dim bs As New BindingSource()
        Assert(bs IsNot Nothing, "BindingSource created successfully")
        Assert(bs.Position = 0, "BindingSource initial position is 0")

        ' ---- Test 2: Create DataTable ----
        Console.WriteLine()
        Console.WriteLine("--- DataTable Creation ---")
        Dim dt As New DataTable("Customers")
        Assert(dt IsNot Nothing, "DataTable created successfully")
        Assert(dt.TableName = "Customers", "DataTable name is Customers")

        ' ---- Test 3: BindingSource.DataSource = DataTable ----
        Console.WriteLine()
        Console.WriteLine("--- BindingSource DataSource ---")
        bs.DataSource = dt
        Assert(bs.DataSource IsNot Nothing, "BindingSource DataSource assigned")

        ' ---- Test 4: BindingSource navigation (empty) ----
        Console.WriteLine()
        Console.WriteLine("--- BindingSource Navigation ---")
        bs.MoveFirst()
        Assert(bs.Position = 0, "MoveFirst on empty stays at 0")
        bs.MoveNext()
        bs.MovePrevious()
        bs.MoveLast()
        Console.WriteLine("PASS: All navigation methods execute without error")
        passed = passed + 1

        ' ---- Test 5: BindingSource DataMember ----
        Console.WriteLine()
        Console.WriteLine("--- BindingSource DataMember ---")
        Assert(bs.DataMember = "", "Initial DataMember is empty")
        bs.DataMember = "Customers"
        Assert(bs.DataMember = "Customers", "DataMember set correctly")

        ' ---- Test 6: BindingSource edit methods ----
        Console.WriteLine()
        Console.WriteLine("--- BindingSource Edit Methods ---")
        bs.EndEdit()
        bs.CancelCurrentEdit()
        bs.ResetBindings()
        Console.WriteLine("PASS: All edit methods execute without error")
        passed = passed + 1

        ' ---- Test 7: Filter and Sort ----
        Console.WriteLine()
        Console.WriteLine("--- BindingSource Filter and Sort ---")
        bs.Filter = "Name = 'Test'"
        Assert(bs.Filter = "Name = 'Test'", "Filter set correctly")
        bs.Sort = "Name ASC"
        Assert(bs.Sort = "Name ASC", "Sort set correctly")

        ' ---- Summary ----
        Console.WriteLine()
        Console.WriteLine("=========================")
        Console.WriteLine("Total: " & (passed + failed).ToString())
        Console.WriteLine("Passed: " & passed.ToString())
        Console.WriteLine("Failed: " & failed.ToString())
        If failed = 0 Then
            Console.WriteLine("ALL TESTS PASSED!")
            Console.WriteLine("SUCCESS")
        End If
    End Sub

End Module
