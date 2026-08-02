// vybe-test: csharp/csharp_tuples_ranges/range_array_slice
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

var arr = new[] { 0, 1, 2, 3, 4 };
var slice = arr[1..4];
foreach (var x in slice) Console.WriteLine(x);
