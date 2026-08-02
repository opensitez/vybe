// vybe-test: csharp/csharp_bitwise_operations/compound_or_assign_sets_bits_in_place
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 0b1000; x |= 0b0011; __Check((x).ToString(), "11");
