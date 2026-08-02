// vybe-test: csharp/csharp_events_advanced/func_delegate_returns_computed_result
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; Func<int, int, int> add = (left, right) => left + right; __Check((add(4, 5)).ToString(), "9");
