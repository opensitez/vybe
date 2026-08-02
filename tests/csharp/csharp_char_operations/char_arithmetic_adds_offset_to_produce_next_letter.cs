// vybe-test: csharp/csharp_char_operations/char_arithmetic_adds_offset_to_produce_next_letter
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char c = (char)('A' + 2); __Check((c).ToString(), "C");
