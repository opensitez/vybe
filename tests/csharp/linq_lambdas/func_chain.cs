// vybe-test: csharp/linq_lambdas/func_chain
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, int> doubleIt = x => x * 2;
Func<int, int> addOne = x => x + 1;
__Check((addOne(doubleIt(5))).ToString(), "11");
__Check((doubleIt(addOne(5))).ToString(), "12");
