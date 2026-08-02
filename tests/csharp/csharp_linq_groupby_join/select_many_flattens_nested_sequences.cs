// vybe-test: csharp/csharp_linq_groupby_join/select_many_flattens_nested_sequences
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

var nested = new[] { new[]{1,2}, new[]{3,4} };
var flat = nested.SelectMany(x => x);
int sum = 0;
foreach (var n in flat) sum += n;
Console.WriteLine(sum);
