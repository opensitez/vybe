// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_string_many_default
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{"a","b"}.SingleOrDefault("z")).ToString(), "z");
