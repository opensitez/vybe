// vybe-test: csharp/csharp_linq_aggregate_element/min_by_largest_number_by_abs
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{-5,2,-1}.MinBy(n=>System.Math.Abs(n))).ToString(), "-1");
