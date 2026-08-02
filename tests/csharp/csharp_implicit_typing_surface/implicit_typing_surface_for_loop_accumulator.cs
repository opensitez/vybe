// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

// implicit_typing_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
