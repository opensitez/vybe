// vybe-test: csharp/csharp_pattern_deconstruct/list_pattern_with_slice_matches_prefix_and_suffix
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

int[] data = { 1, 2, 3, 4, 5 };
if (data is [1, .., 5]) Console.WriteLine("bookended");
else Console.WriteLine("no");
