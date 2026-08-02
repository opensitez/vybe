// vybe-test: csharp/csharp_bitwise_operations/compound_bitwise_and_assign_updates_in_place
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 0b1111; x &= 0b0101; __Check((x).ToString(), "5");
