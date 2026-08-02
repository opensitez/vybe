// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

// implicit_typing_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
