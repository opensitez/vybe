// vybe-test: csharp/linq_lambdas/predicate_usage
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Predicate<int> isEven = x => x % 2 == 0;
__Check((isEven(4)).ToString(), "True");
__Check((isEven(7)).ToString(), "False");
