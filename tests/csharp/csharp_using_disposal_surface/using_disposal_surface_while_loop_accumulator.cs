// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

// using_disposal_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
