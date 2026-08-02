// vybe-test: csharp/csharp_with_expression_records/with_string_empty
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Label(string T); var e=(new Label("a")) with{T=""}; __Check((e.T.Length).ToString(), "0");
