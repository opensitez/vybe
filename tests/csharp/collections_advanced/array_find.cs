// vybe-test: csharp/collections_advanced/array_find
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = { 10, 20, 30, 40 };
__Check((Array.Find(arr, x => x > 15)).ToString(), "20");
