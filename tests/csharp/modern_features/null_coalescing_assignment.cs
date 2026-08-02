// vybe-test: csharp/modern_features/null_coalescing_assignment
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
s ??= "assigned";
__Check((s).ToString(), "assigned");
s ??= "not again";
__Check((s).ToString(), "assigned");
