// vybe-test: csharp/csharp_comparison_sorting/default_comparer_orders_smaller_int_before_larger_int
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Collections.Generic.Comparer<int>.Default.Compare(2, 5)).ToString(), "-1");
