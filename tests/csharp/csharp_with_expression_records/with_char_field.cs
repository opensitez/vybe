// vybe-test: csharp/csharp_with_expression_records/with_char_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Sym(char C); var b=(new Sym('a')) with{C='z'}; __Check((b.C).ToString(), "z");
