// vybe-test: csharp/csharp_linq_numeric/sum_of_empty_sequence_is_zero
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Array.Empty<int>().Sum()).ToString(), "0");
