// vybe-test: csharp/linq_lambdas/linq_orderby_thenby
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

var names = new List<string> { "Charlie", "Alice", "Bob", "Alice" };
var sorted = names.OrderBy(n => n).ToList();
foreach (var n in sorted) Console.WriteLine(n);
