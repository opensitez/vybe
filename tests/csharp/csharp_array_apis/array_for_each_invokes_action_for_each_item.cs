// vybe-test: csharp/csharp_array_apis/array_for_each_invokes_action_for_each_item
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

var values = new[] { 3, 4 }; System.Array.ForEach(values, value => Console.WriteLine(value * 2));
