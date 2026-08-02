// vybe-test: csharp/csharp_linq_aggregates/count_with_predicate_counts_matching_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3,4,5}.Count(x => x%2==0)).ToString(), "2");
