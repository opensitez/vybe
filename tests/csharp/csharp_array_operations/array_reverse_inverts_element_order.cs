// vybe-test: csharp/csharp_array_operations/array_reverse_inverts_element_order
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = {1,2,3};
System.Array.Reverse(a);
__Check((a[0]).ToString(), "3");
