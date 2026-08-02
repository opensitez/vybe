// vybe-test: csharp/csharp_nullable/null_in_ternary
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
__Check((s != null ? s : "none").ToString(), "none");
s = "found";
__Check((s != null ? s : "none").ToString(), "found");
