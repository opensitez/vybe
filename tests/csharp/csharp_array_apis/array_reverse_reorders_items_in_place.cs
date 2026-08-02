// vybe-test: csharp/csharp_array_apis/array_reverse_reorders_items_in_place
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

var values = new[] { 1, 2, 3 }; System.Array.Reverse(values); foreach (var value in values) Console.WriteLine(value);
