// vybe-test: csharp/csharp_events_advanced/predicate_delegate_filters_positive_values
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; Predicate<int> positive = value => value > 0; __Check((positive(3)).ToString(), "True"); __Check((positive(-1)).ToString(), "False");
