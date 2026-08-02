// vybe-test: csharp/collections_advanced/array_indexof
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] arr = { "a", "b", "c", "d" };
__Check((Array.IndexOf(arr, "c")).ToString(), "2");
__Check((Array.IndexOf(arr, "z")).ToString(), "-1");
