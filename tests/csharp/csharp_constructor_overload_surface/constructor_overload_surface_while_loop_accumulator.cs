// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

// constructor_overload_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
