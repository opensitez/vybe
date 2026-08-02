// vybe-test: csharp/csharp_array_operations/array_clear_fills_range_with_default_values
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = {1,2,3,4,5};
System.Array.Clear(a, 1, 3);
__Check((a[0]).ToString(), "1"); __Check((a[2]).ToString(), "0");
