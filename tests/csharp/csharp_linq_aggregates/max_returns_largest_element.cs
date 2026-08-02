// vybe-test: csharp/csharp_linq_aggregates/max_returns_largest_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{5,1,9,3}.Max()).ToString(), "9");
