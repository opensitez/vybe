// vybe-test: csharp/csharp_linq_aggregate_element/element_at_or_default_in_range
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{5,6,7}.ElementAtOrDefault(1)).ToString(), "6");
