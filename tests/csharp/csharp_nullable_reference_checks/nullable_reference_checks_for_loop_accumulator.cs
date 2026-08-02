// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

// nullable_reference_checks
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
