// vybe-test: csharp/csharp_array_apis/array_clear_resets_values_to_default
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

var values = new[] { 1, 2, 3 }; System.Array.Clear(values, 1, 2); foreach (var value in values) Console.WriteLine(value);
