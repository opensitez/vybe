// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

// string_immutability_checks
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
