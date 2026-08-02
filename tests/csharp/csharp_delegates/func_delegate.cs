// vybe-test: csharp/csharp_delegates/func_delegate
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, int> square = x => x * x;
__Check((square(5)).ToString(), "25");
__Check((square(8)).ToString(), "64");
