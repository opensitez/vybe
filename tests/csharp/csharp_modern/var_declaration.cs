// vybe-test: csharp/csharp_modern/var_declaration
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 42;
var s = "hello";
__Check((x).ToString(), "42");
__Check((s).ToString(), "hello");
