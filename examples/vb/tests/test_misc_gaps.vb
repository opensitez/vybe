Imports System
Imports System.Collections.Generic

Module TestMiscGaps
    Class TestPerson
        Public Name As String = "Alice"
        Public Age As Integer = 30

        Public Function SayHello() As String
            return "Hello " & Name
        End Function
        
        Public Sub New()
            Name = "Constructor"
        End Sub
        
        Public Sub Rename(newName As String)
            Name = newName
        End Sub
    End Class

    Sub Main()
        TestCallByName()
        TestActivator()
        TestDictionary()
    End Sub

    Sub TestCallByName()
        Console.WriteLine("--- TestCallByName ---")
        Dim p As New TestPerson()

        ' Get
        Dim n = CallByName(p, "Name", 2)
        AssertEqual(n, "Constructor")
        Console.WriteLine("Get Name: " & n)

        ' Set
        CallByName(p, "Name", 4, "Bob")
        AssertEqual(p.Name, "Bob")
        Console.WriteLine("Set Name to Bob. p.Name=" & p.Name)

        ' Method (CallType=1)
        ' Note: My implementation currently returns Nothing for method call in CallByName case 1? 
        ' Let's check if I implemented it. I returned Nothing in default case.
        ' So this might fail if I expect return value.
        ' If CallByName is used as function for method call, it should return value.
        
        ' CallByName(p, "SayHello", 1) -> Should return string?
        Dim s = CallByName(p, "SayHello", 1)
        AssertEqual(s, "Hello Bob")
        Console.WriteLine("Method result: " & s)
    End Sub

    Sub TestActivator()
        Console.WriteLine("--- TestActivator ---")
        ' Dim obj = Activator.CreateInstance("TestPerson")
        ' Use late binding or direct cast if possible?
        ' For now, just check if it returns object.
        
        Dim obj As Object = Activator.CreateInstance("TestPerson")
        If obj Is Nothing Then
            Console.WriteLine("FAIL: Activator returned Nothing")
        Else
            Console.WriteLine("PASS: Activator returned object")
            ' Check if default values are set (Alice, 30)
            ' Using CallByName to check properties on late-bound object
            Dim n = CallByName(obj, "Name", 2)
            AssertEqual(n, "Constructor")
            
            ' Verify if constructor was called.
            ' But TestPerson doesn't have custom constructor affecting state currently (fields are initialized).
            ' Let's add a constructor effect?
            ' I can't modify class easily inside Sub.
            ' But fields with initializers SHOULD be set by Activator implementation.
        End If
    End Sub

    Sub TestDictionary()
        Console.WriteLine("--- TestDictionary ---")
        Dim d As New Dictionary(Of String, Integer)()
        d.Add("One", 1)
        AssertEqual(d.Count, 1)
        AssertEqual(d("One"), 1)
        
        If d.ContainsKey("One") Then
            Console.WriteLine("ContainsKey works")
        End If
    End Sub

    Sub AssertEqual(actual As Object, expected As Object)
        If actual <> expected Then
            Console.WriteLine("FAIL: Expected " & expected & ", but got " & actual)
        Else
            Console.WriteLine("PASS")
        End If
    End Sub
End Module
