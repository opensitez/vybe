// vybe-test: csharp/csharp_linq_aggregates/all_returns_false_when_one_element_fails_predicate
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{2,4,5}.All(x => x%2==0)).ToString(), "False");
