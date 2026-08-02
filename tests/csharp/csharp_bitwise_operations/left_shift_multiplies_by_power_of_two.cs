// vybe-test: csharp/csharp_bitwise_operations/left_shift_multiplies_by_power_of_two
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((1 << 4).ToString(), "16");
