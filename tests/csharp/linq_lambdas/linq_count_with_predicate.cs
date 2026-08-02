// vybe-test: csharp/linq_lambdas/linq_count_with_predicate
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int> { 1, 2, 3, 4, 5, 6 };
__Check((nums.Count(x => x > 3)).ToString(), "3");
