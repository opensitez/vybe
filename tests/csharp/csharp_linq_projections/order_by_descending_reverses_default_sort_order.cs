// vybe-test: csharp/csharp_linq_projections/order_by_descending_reverses_default_sort_order
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

var result = new[]{3,1,4,1,5}.OrderByDescending(x => x).Distinct();
foreach(var n in result) Console.WriteLine(n);
