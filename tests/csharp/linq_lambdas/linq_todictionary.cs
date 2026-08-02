// vybe-test: csharp/linq_lambdas/linq_todictionary
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var names = new List<string> { "Alice", "Bob" };
var dict = names.ToDictionary(n => n, n => n.Length);
__Check((dict["Alice"]).ToString(), "5");
__Check((dict["Bob"]).ToString(), "3");
