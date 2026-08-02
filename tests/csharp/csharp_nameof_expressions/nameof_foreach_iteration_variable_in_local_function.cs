// vybe-test: csharp/csharp_nameof_expressions/nameof_foreach_iteration_variable_in_local_function
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void Scan(){foreach(var entry in new string[]{"a"}){Console.WriteLine(nameof(entry)); break;}} Scan();
