// vybe-test: csharp/linq_lambdas/linq_where_select_tolist
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

var nums = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
var result = nums.Where(x => x % 2 == 0).Select(x => x * 10).ToList();
foreach (var x in result) Console.WriteLine(x);
