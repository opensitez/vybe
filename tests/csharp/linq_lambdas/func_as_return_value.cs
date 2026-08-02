// vybe-test: csharp/linq_lambdas/func_as_return_value
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, Func<int, int>> makeAdder = x => y => x + y;
var add5 = makeAdder(5);
__Check((add5(3)).ToString(), "8");
__Check((add5(10)).ToString(), "15");
