// vybe-test: csharp/csharp_nullable/null_coalescing_assign
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
s ??= "assigned";
__Check((s).ToString(), "assigned");
s ??= "not this";
__Check((s).ToString(), "assigned");
