// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

// exception_type_checks
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
