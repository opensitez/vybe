// vybe-test: csharp/strings_advanced/string_compareto
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a = "apple";
string b = "banana";
__Check((a.CompareTo(b) < 0).ToString(), "True");
__Check((b.CompareTo(a) > 0).ToString(), "True");
__Check((a.CompareTo(a) == 0).ToString(), "True");
