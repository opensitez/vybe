// vybe-test: csharp/csharp_throw_in_null_coalescing/chained_null_coalescing_throw_only_evaluates_when_all_prior_operands_null
// origin: languages/csharp/tests/csharp/test_csharp_throw_in_null_coalescing.rs

string? a = null;
string? b = null;
try {
    string value = a ?? b ?? throw new System.Exception("both-null");
    Console.WriteLine(value);
} catch (System.Exception) {
    Console.WriteLine("caught");
}
