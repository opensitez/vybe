// vybe-test: csharp/linq_lambdas/linq_aggregate
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int> { 1, 2, 3, 4, 5 };
var product = nums.Aggregate(1, (acc, x) => acc * x);
__Check((product).ToString(), "120");
