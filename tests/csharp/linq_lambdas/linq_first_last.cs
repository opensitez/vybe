// vybe-test: csharp/linq_lambdas/linq_first_last
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int> { 10, 20, 30 };
__Check((nums.First()).ToString(), "10");
__Check((nums.Last()).ToString(), "30");
