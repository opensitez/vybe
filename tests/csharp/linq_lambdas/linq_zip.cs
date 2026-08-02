// vybe-test: csharp/linq_lambdas/linq_zip
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

var names = new List<string> { "Alice", "Bob", "Charlie" };
var ages = new List<int> { 30, 25, 35 };
var pairs = names.Zip(ages, (n, a) => n + "=" + a).ToList();
foreach (var p in pairs) Console.WriteLine(p);
