// vybe-test: csharp/csharp_pattern_matching_advanced/switch_statement_type_pattern_matches_int_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

object item = 9; switch (item) { case string text: Console.WriteLine(text); break; case int number: Console.WriteLine(number * 3); break; default: Console.WriteLine("other"); break; }
