// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

// exception_type_checks
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
