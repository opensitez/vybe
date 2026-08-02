// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

// string_immutability_checks
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
