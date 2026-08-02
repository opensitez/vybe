// vybe-test: csharp/csharp_operators/null_coalescing
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

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
