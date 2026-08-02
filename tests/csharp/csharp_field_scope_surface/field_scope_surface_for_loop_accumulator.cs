// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

// field_scope_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
