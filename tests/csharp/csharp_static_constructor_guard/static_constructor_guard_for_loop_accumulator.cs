// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

// static_constructor_guard
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
