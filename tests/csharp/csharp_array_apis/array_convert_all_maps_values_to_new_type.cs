// vybe-test: csharp/csharp_array_apis/array_convert_all_maps_values_to_new_type
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

var values = new[] { 1, 2, 3 }; var text = System.Array.ConvertAll(values, value => "n" + value); foreach (var value in text) Console.WriteLine(value);
