// vybe-test: csharp/modern_features/null_conditional_operator
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
__Check((s?.ToUpper() ?? "null").ToString(), "null");
s = "hello";
__Check((s?.ToUpper() ?? "null").ToString(), "HELLO");
