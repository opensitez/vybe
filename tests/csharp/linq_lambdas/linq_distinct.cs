// vybe-test: csharp/linq_lambdas/linq_distinct
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int> { 1, 2, 2, 3, 3, 3, 4 };
var distinct = nums.Distinct().ToList();
__Check((distinct.Count).ToString(), "4");
