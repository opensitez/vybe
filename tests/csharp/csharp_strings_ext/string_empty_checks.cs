// vybe-test: csharp/csharp_strings_ext/string_empty_checks
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "";
__Check((s.Length).ToString(), "0");
__Check((s == "").ToString(), "True");
