// vybe-test: csharp/linq_runtime/linq_any_pred
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
__Check((list.Any(x => x > 10)).ToString(), "False");
__Check((list.Any(x => x == 2)).ToString(), "True");
