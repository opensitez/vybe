// vybe-test: csharp/csharp_linq_aggregates/first_with_predicate_skips_to_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3,4}.First(x => x>2)).ToString(), "3");
