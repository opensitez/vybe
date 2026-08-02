// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

// using_disposal_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
