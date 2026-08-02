// vybe-test: csharp/csharp_array_operations/array_resize_grows_array_and_preserves_existing_elements
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = {1,2,3};
System.Array.Resize(ref a, 5);
__Check((a.Length).ToString(), "5"); __Check((a[2]).ToString(), "3");
