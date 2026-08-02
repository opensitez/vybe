// vybe-test: csharp/csharp_bitwise_operations/bitwise_not_inverts_all_bits_of_byte
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte b = 0b11110000; __Check(((byte)(~b)).ToString(), "15");
