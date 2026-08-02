// vybe-test: csharp/csharp_array_apis/instance_copy_to_moves_values_into_target_array
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

var source = new[] { 9, 8 }; var target = new int[2]; source.CopyTo(target, 0); foreach (var value in target) Console.WriteLine(value);
