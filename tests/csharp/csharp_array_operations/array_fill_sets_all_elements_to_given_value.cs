// vybe-test: csharp/csharp_array_operations/array_fill_sets_all_elements_to_given_value
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = new int[4];
System.Array.Fill(a, 7);
__Check((a[3]).ToString(), "7");
