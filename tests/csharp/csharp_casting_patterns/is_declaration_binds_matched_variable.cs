// vybe-test: csharp/csharp_casting_patterns/is_declaration_binds_matched_variable
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=42;
if(o is int n) __Check((n*2).ToString(), "84");
