// vybe-test: csharp/csharp_modern/boolean_values
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool t = true;
bool f = false;
__Check((t).ToString(), "True");
__Check((f).ToString(), "False");
__Check((t && f).ToString(), "False");
__Check((t || f).ToString(), "True");
__Check((!t).ToString(), "False");
