// vybe-test: csharp/csharp_array_apis/array_copy_moves_values_between_arrays
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

var source = new[] { 5, 6, 7 }; var target = new int[3]; System.Array.Copy(source, target, 3); foreach (var value in target) Console.WriteLine(value);
