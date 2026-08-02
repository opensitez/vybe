// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

// static_constructor_guard
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
