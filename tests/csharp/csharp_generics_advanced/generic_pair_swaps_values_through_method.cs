// vybe-test: csharp/csharp_generics_advanced/generic_pair_swaps_values_through_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(T, T) Swap<T>(T a, T b) => (b, a);
var (x, y) = Swap(1, 2);
__Check((x).ToString(), "2"); __Check((y).ToString(), "1");
