// vybe-test: csharp/linq_lambdas/linq_contains
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int> { 1, 2, 3, 4 };
__Check((nums.Contains(3)).ToString(), "True");
__Check((nums.Contains(9)).ToString(), "False");
