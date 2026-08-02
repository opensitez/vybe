// vybe-test: csharp/csharp_linq_aggregates/contains_returns_true_for_present_value_in_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3}.Contains(2)).ToString(), "True");
