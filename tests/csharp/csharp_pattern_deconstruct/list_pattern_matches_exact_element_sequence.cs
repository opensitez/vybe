// vybe-test: csharp/csharp_pattern_deconstruct/list_pattern_matches_exact_element_sequence
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

int[] data = { 1, 2, 3 };
if (data is [1, 2, 3]) Console.WriteLine("exact");
else Console.WriteLine("no");
