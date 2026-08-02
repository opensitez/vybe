// vybe-test: csharp/linq_lambdas/linq_selectmany
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

var lists = new List<List<int>> {
    new List<int> { 1, 2 },
    new List<int> { 3, 4 },
    new List<int> { 5 }
};
var flat = lists.SelectMany(l => l).ToList();
foreach (var x in flat) Console.WriteLine(x);
