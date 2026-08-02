// vybe-test: csharp/csharp_array_apis/array_resize_grows_array_and_preserves_existing_values
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

var values = new[] { 2, 4 }; System.Array.Resize(ref values, 4); foreach (var value in values) Console.WriteLine(value);
