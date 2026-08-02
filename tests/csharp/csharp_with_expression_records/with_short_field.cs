// vybe-test: csharp/csharp_with_expression_records/with_short_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record ShortBox(short S); var t=(new ShortBox(1)) with{S=1000}; __Check((t.S).ToString(), "1000");
