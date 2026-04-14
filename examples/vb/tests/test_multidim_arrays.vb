Imports System

Module TestMultiDimArrays
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, name As String)
        If condition Then
            passed = passed + 1
            Console.WriteLine("  PASS: " & name)
        Else
            failed = failed + 1
            Console.WriteLine("  FAIL: " & name)
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== Multi-Dimensional Array Tests ===")

        ' --- 2D array declaration ---
        Console.WriteLine("Test: 2D array declaration")
        Dim grid(2, 3) As Integer
        Assert(True, "2D array declared without error")

        ' --- 2D array assignment and read-back ---
        Console.WriteLine("Test: 2D array assignment")
        grid(0, 0) = 1
        grid(0, 1) = 2
        grid(0, 2) = 3
        grid(1, 0) = 4
        grid(1, 1) = 5
        grid(1, 2) = 6
        grid(2, 0) = 7
        grid(2, 1) = 8
        grid(2, 2) = 9
        Assert(grid(0, 0) = 1, "2D: (0,0) = 1")
        Assert(grid(0, 2) = 3, "2D: (0,2) = 3")
        Assert(grid(1, 1) = 5, "2D: (1,1) = 5")
        Assert(grid(2, 2) = 9, "2D: (2,2) = 9")

        ' --- 2D array overwrite ---
        Console.WriteLine("Test: 2D array overwrite")
        grid(1, 1) = 99
        Assert(grid(1, 1) = 99, "2D: overwrite (1,1) = 99")
        Assert(grid(1, 0) = 4, "2D: neighbor unchanged")

        ' --- 3D array ---
        Console.WriteLine("Test: 3D array declaration")
        Dim cube(1, 1, 1) As Integer
        cube(0, 0, 0) = 100
        cube(0, 0, 1) = 200
        cube(0, 1, 0) = 300
        cube(1, 0, 0) = 400
        cube(1, 1, 1) = 500
        Assert(cube(0, 0, 0) = 100, "3D: (0,0,0) = 100")
        Assert(cube(0, 0, 1) = 200, "3D: (0,0,1) = 200")
        Assert(cube(0, 1, 0) = 300, "3D: (0,1,0) = 300")
        Assert(cube(1, 0, 0) = 400, "3D: (1,0,0) = 400")
        Assert(cube(1, 1, 1) = 500, "3D: (1,1,1) = 500")

        ' --- 2D default values ---
        Console.WriteLine("Test: 2D default values")
        Dim defaults(1, 1) As Integer
        Assert(defaults(0, 0) = 0, "2D default: (0,0) = 0")
        Assert(defaults(1, 1) = 0, "2D default: (1,1) = 0")

        ' --- 1D still works ---
        Console.WriteLine("Test: 1D array still works")
        Dim arr(4) As Integer
        arr(0) = 10
        arr(4) = 50
        Assert(arr(0) = 10, "1D: index 0 = 10")
        Assert(arr(4) = 50, "1D: index 4 = 50")

        Console.WriteLine("")
        Console.WriteLine("Results: " & passed & " passed, " & failed & " failed")
        If failed = 0 Then
            Console.WriteLine("=== ALL MULTI-DIM ARRAY TESTS PASSED ===")
        Else
            Console.WriteLine("=== SOME MULTI-DIM ARRAY TESTS FAILED ===")
        End If
    End Sub
End Module
