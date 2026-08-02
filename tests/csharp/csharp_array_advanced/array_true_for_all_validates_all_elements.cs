// vybe-test: csharp/csharp_array_advanced/array_true_for_all_validates_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={2,4,6,8};
__Check((System.Array.TrueForAll(arr,n=>n%2==0)).ToString(), "True");
