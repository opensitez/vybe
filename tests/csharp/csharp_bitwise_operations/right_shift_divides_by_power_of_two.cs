// vybe-test: csharp/csharp_bitwise_operations/right_shift_divides_by_power_of_two
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((64 >> 3).ToString(), "8");
