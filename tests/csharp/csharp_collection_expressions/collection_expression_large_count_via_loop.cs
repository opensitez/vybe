// vybe-test: csharp/csharp_collection_expressions/collection_expression_large_count_via_loop
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

int[] arr = [1, 2, 3, 4, 5];
int sum = 0;
for (int i = 0; i < arr.Length; i++) sum += arr[i];
Console.WriteLine(sum);
