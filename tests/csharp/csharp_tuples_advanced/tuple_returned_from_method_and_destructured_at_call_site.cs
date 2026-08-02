// vybe-test: csharp/csharp_tuples_advanced/tuple_returned_from_method_and_destructured_at_call_site
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(int Min, int Max) Bounds(int[] arr) =>
    (arr.Min(), arr.Max());
var (lo, hi) = Bounds(new[]{5,1,9,3});
__Check((lo).ToString(), "1"); __Check((hi).ToString(), "9");
