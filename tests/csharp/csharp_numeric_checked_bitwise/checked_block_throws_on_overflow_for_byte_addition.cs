// vybe-test: csharp/csharp_numeric_checked_bitwise/checked_block_throws_on_overflow_for_byte_addition
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

try { checked { byte value = 255; value += 1; } Console.WriteLine("no-throw"); } catch (System.OverflowException) { Console.WriteLine("overflow"); }
