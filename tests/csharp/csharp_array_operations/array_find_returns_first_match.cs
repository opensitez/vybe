// vybe-test: csharp/csharp_array_operations/array_find_returns_first_match
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = {1,3,5,7};
__Check((System.Array.Find(a, x => x > 3)).ToString(), "5");
