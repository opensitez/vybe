// vybe-test: csharp/csharp_icomparable_sorting/linq_order_by_uses_default_icomparable_for_value_types
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

var result = new[]{3,1,2}.OrderBy(x=>x);
foreach(var n in result) Console.WriteLine(n);
