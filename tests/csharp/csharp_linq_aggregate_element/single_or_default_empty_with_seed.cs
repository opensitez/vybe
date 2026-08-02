// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_empty_with_seed
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Array.Empty<int>().SingleOrDefault(99)).ToString(), "99");
