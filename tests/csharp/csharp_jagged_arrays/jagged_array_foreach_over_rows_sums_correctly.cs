// vybe-test: csharp/csharp_jagged_arrays/jagged_array_foreach_over_rows_sums_correctly
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

int[][] jag = new[]{ new[]{1,2}, new[]{3,4,5} };
int total=0;
foreach(var row in jag) foreach(var v in row) total+=v;
Console.WriteLine(total);
