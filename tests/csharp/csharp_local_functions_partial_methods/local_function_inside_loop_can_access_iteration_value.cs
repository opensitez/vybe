// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_inside_loop_can_access_iteration_value
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

foreach (var item in new[] { 1, 2, 3 }) { int Square() { return item * item; } Console.WriteLine(Square()); }
