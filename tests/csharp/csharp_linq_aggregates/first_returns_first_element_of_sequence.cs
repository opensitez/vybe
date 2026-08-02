// vybe-test: csharp/csharp_linq_aggregates/first_returns_first_element_of_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{10,20,30}.First()).ToString(), "10");
