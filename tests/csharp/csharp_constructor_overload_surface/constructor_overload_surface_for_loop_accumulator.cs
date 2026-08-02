// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

// constructor_overload_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
