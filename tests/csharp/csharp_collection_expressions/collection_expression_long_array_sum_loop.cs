// vybe-test: csharp/csharp_collection_expressions/collection_expression_long_array_sum_loop
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

long[] nums = [10000000000L, 20000000000L];
long total = 0;
foreach (var n in nums) total += n;
Console.WriteLine(total);
