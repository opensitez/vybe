// vybe-test: csharp/csharp_delegates/func_two_args
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, int, int> add = (a, b) => a + b;
__Check((add(3, 4)).ToString(), "7");
