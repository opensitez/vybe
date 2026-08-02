// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

// pattern_constant_checks
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
