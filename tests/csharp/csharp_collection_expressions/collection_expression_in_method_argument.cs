// vybe-test: csharp/csharp_collection_expressions/collection_expression_in_method_argument
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

int Sum(int[] data) { int t = 0; foreach (var n in data) t += n; return t; }
Console.WriteLine(Sum([1, 2, 3]));
