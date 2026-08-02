// vybe-test: csharp/csharp_throw_in_null_coalescing/null_coalescing_throw_expression_runs_when_left_is_null
// origin: languages/csharp/tests/csharp/test_csharp_throw_in_null_coalescing.rs

string? missing = null;
try {
    string value = missing ?? throw new System.Exception("required");
    Console.WriteLine(value);
} catch (System.Exception e) {
    Console.WriteLine(e.Message);
}
