// vybe-test: csharp/csharp_nullable/null_coalescing_string
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
__Check((s ?? "fallback").ToString(), "fallback");
s = "hello";
__Check((s ?? "fallback").ToString(), "hello");
