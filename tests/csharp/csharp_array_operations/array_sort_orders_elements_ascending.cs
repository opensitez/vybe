// vybe-test: csharp/csharp_array_operations/array_sort_orders_elements_ascending
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = {3,1,4,1,5};
System.Array.Sort(a);
__Check((a[0]).ToString(), "1"); __Check((a[4]).ToString(), "5");
