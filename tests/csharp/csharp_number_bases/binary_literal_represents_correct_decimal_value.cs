// vybe-test: csharp/csharp_number_bases/binary_literal_represents_correct_decimal_value
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0b1010).ToString(), "10");
