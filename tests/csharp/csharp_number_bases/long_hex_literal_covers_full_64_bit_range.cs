// vybe-test: csharp/csharp_number_bases/long_hex_literal_covers_full_64_bit_range
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0x7FFFFFFFFFFFFFFFL==long.MaxValue).ToString(), "True");
