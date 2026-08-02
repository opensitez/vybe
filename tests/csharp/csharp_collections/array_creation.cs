// vybe-test: csharp/csharp_collections/array_creation
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = {5, 10, 15, 20, 25};
__Check((arr.Length).ToString(), "5");
__Check((arr[2]).ToString(), "15");
