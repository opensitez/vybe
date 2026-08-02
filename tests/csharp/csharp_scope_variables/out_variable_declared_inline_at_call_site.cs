// vybe-test: csharp/csharp_scope_variables/out_variable_declared_inline_at_call_site
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

if(int.TryParse("42", out int n)) Console.WriteLine(n);
else Console.WriteLine(0);
