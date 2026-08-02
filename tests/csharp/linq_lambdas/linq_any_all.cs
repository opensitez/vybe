// vybe-test: csharp/linq_lambdas/linq_any_all
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int> { 2, 4, 6, 8 };
__Check((nums.All(x => x % 2 == 0)).ToString(), "True");
__Check((nums.Any(x => x > 5)).ToString(), "True");
__Check((nums.Any(x => x > 10)).ToString(), "False");
