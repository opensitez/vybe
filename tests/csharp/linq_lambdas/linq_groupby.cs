// vybe-test: csharp/linq_lambdas/linq_groupby
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

var words = new List<string> { "apple", "ant", "banana", "avocado", "bat" };
var groups = words.GroupBy(w => w[0].ToString()).ToList();
foreach (var g in groups) {
    Console.WriteLine(g.Key + ": " + g.Count());
}
