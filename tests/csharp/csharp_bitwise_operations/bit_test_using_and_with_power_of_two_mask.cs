// vybe-test: csharp/csharp_bitwise_operations/bit_test_using_and_with_power_of_two_mask
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int flags = 0b1010; __Check(((flags & 0b0010) != 0).ToString(), "True");
