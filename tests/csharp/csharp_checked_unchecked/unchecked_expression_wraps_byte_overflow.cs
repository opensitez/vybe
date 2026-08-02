// vybe-test: csharp/csharp_checked_unchecked/unchecked_expression_wraps_byte_overflow
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte b=unchecked((byte)256); __Check((b).ToString(), "0");
