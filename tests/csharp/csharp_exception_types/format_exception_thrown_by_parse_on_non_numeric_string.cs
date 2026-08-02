// vybe-test: csharp/csharp_exception_types/format_exception_thrown_by_parse_on_non_numeric_string
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "";
try { int.Parse("abc"); }
catch(System.FormatException) { result = "fmt"; }
__Check((result).ToString(), "fmt");
