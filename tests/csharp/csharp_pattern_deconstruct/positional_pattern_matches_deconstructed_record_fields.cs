// vybe-test: csharp/csharp_pattern_deconstruct/positional_pattern_matches_deconstructed_record_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

record Point(int X, int Y);
object obj = new Point(0, 5);
if (obj is Point(0, var y)) Console.WriteLine(y);
else Console.WriteLine(-1);
