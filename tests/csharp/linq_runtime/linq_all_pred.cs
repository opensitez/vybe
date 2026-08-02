// vybe-test: csharp/linq_runtime/linq_all_pred
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
__Check((list.All(x => x > 0)).ToString(), "True");
__Check((list.All(x => x > 2)).ToString(), "False");
