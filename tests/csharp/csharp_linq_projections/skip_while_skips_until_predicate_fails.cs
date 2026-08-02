// vybe-test: csharp/csharp_linq_projections/skip_while_skips_until_predicate_fails
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

var result = new[]{1,2,3,4,5}.SkipWhile(x => x<3);
foreach(var n in result) Console.WriteLine(n);
