// vybe-test: csharp/linq_lambdas/linq_orderbydescending
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

var nums = new List<int> { 3, 1, 4, 1, 5 };
var sorted = nums.OrderByDescending(x => x).ToList();
foreach (var x in sorted) Console.WriteLine(x);
