// vybe-test: csharp/modern_features/null_coalescing_operator
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
__Check((s ?? "default").ToString(), "default");
s = "hello";
__Check((s ?? "default").ToString(), "hello");
