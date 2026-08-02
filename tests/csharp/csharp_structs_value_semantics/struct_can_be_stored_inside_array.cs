// vybe-test: csharp/csharp_structs_value_semantics/struct_can_be_stored_inside_array
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

struct Point { public int X; } var points = new[] { new Point { X = 3 }, new Point { X = 4 } }; foreach (var point in points) Console.WriteLine(point.X);
