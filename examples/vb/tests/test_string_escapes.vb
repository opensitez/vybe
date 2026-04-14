Module Program
    Sub Main()
        ' Test 1: JSON with escaped quotes
        Dim json As String = "{""name"":""vybe"",""version"":""1.0""}"
        Console.WriteLine(json)

        ' Test 2: Empty string
        Dim empty As String = ""
        Console.WriteLine("Empty: [" & empty & "]")

        ' Test 3: Single embedded quote
        Dim oneQuote As String = """"
        Console.WriteLine("One quote: [" & oneQuote & "]")

        ' Test 4: Embedded quotes in sentence
        Dim msg As String = "She said ""hello"" to me"
        Console.WriteLine(msg)

        ' Test 5: Multiple escaped quotes in a row
        Dim multi As String = "A""""B"
        Console.WriteLine("Multi: [" & multi & "]")

        ' Test 6: Escaped quote at start and end
        Dim edges As String = """hello"""
        Console.WriteLine("Edges: [" & edges & "]")

        ' Test 7: String with only escaped quotes
        Dim allQuotes As String = """"""
        Console.WriteLine("All quotes: [" & allQuotes & "]")
    End Sub
End Module
