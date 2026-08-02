// vybe-test: csharp/csharp_nullable/null_conditional_member
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Wrapper {
    public string Value;
    public Wrapper(string v) { Value = v; }
}
Wrapper w = null;
__Check((w?.Value ?? "null").ToString(), "null");
w = new Wrapper("hello");
__Check((w?.Value ?? "null").ToString(), "hello");
