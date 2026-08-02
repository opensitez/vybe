// vybe-test: csharp/csharp_pattern_matching/type_pattern_matches_int_in_if_is_expression
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

object o = 5; if(o is int n) Console.WriteLine(n); else Console.WriteLine(0);
