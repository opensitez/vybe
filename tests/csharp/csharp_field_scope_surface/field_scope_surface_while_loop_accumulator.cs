// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

// field_scope_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
