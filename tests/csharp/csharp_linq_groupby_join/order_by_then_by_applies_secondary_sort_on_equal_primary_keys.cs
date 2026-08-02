// vybe-test: csharp/csharp_linq_groupby_join/order_by_then_by_applies_secondary_sort_on_equal_primary_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

var items = new[] { (Name:"b",Age:2),(Name:"a",Age:3),(Name:"a",Age:1) };
var sorted = items.OrderBy(x => x.Name).ThenBy(x => x.Age);
foreach (var x in sorted) Console.WriteLine($"{x.Name}{x.Age}");
