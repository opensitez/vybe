// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_empty_returns_default
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Array.Empty<int>().SingleOrDefault()).ToString(), "0");
