// vybe-test: csharp/csharp_pattern_matching_advanced/switch_statement_when_guard_matches_large_number
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

var x = 12; switch (x) { case int number when number > 10: Console.WriteLine("large"); break; case int number: Console.WriteLine("small"); break; }
