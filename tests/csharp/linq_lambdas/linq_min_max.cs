// vybe-test: csharp/linq_lambdas/linq_min_max
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int> { 5, 3, 8, 1, 9 };
__Check((nums.Min()).ToString(), "1");
__Check((nums.Max()).ToString(), "9");
