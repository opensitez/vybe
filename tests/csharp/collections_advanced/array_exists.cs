// vybe-test: csharp/collections_advanced/array_exists
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = { 1, 2, 3, 4, 5 };
__Check((Array.Exists(arr, x => x > 4)).ToString(), "True");
__Check((Array.Exists(arr, x => x > 10)).ToString(), "False");
