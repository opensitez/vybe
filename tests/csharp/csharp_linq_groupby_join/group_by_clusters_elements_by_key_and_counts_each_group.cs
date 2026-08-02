// vybe-test: csharp/csharp_linq_groupby_join/group_by_clusters_elements_by_key_and_counts_each_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

var words = new[] { "apple", "ant", "banana", "bear", "avocado" };
var groups = words
    .GroupBy(w => w[0])
    .OrderBy(g => g.Key)
    .Select(g => $"{g.Key}:{g.Count()}");
foreach (var s in groups) Console.WriteLine(s);
