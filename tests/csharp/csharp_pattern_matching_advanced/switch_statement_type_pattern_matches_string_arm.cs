// vybe-test: csharp/csharp_pattern_matching_advanced/switch_statement_type_pattern_matches_string_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

object item = "beta"; switch (item) { case string text: Console.WriteLine(text.ToUpper()); break; default: Console.WriteLine("other"); break; }
