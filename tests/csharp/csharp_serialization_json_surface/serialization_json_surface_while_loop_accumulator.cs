// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

// serialization_json_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
