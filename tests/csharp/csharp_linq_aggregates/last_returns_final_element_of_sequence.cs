// vybe-test: csharp/csharp_linq_aggregates/last_returns_final_element_of_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{10,20,30}.Last()).ToString(), "30");
