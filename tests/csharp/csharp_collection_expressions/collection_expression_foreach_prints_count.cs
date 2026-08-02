// vybe-test: csharp/csharp_collection_expressions/collection_expression_foreach_prints_count
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

int[] arr = [1, 2, 3];
int count = 0;
foreach (var _ in arr) count++;
Console.WriteLine(count);
