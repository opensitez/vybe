// vybe-test: csharp/csharp_nullable_ref_types/null_forgiving_operator_suppresses_warning_still_nulls_at_runtime
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

string? s=null;
string r="ok";
try{Console.WriteLine(s!.Length);}
catch(System.NullReferenceException){r="null";}
Console.WriteLine(r);
