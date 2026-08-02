// vybe-test: csharp/csharp_array_operations/array_exists_detects_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = {1,3,5,7};
__Check((System.Array.Exists(a, x => x > 4)).ToString(), "True");
