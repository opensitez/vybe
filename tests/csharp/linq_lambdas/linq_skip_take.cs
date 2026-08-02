// vybe-test: csharp/linq_lambdas/linq_skip_take
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

var nums = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
var page = nums.Skip(2).Take(3).ToList();
foreach (var x in page) Console.WriteLine(x);
