// vybe-test: csharp/collections_advanced/array_reverse
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

int[] arr = { 1, 2, 3, 4, 5 };
Array.Reverse(arr);
foreach (var x in arr) Console.WriteLine(x);
