// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

// explicit_typing_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
