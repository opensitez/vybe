// vybe-test: csharp/collections_advanced/array_initialization_syntax
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = { 1, 2, 3, 4, 5 };
__Check((a.Length).ToString(), "5");
__Check((a[0] + a[4]).ToString(), "6");
