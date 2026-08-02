// vybe-test: csharp/csharp_linq_aggregates/any_returns_true_when_predicate_satisfied
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3}.Any(x => x>2)).ToString(), "True");
