// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

// linq_query_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
