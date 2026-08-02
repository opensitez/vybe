// vybe-test: csharp/csharp_events_advanced/delegate_return_value_can_be_consumed_by_caller
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Calculator { public static int Compute(Func<int> getValue) { return getValue() + 1; } } __Check((Calculator.Compute(() => 9)).ToString(), "10");
