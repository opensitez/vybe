// vybe-test: csharp/csharp_linq_aggregates/min_returns_smallest_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{5,1,9,3}.Min()).ToString(), "1");
