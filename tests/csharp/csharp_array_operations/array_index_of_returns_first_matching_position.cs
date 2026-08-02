// vybe-test: csharp/csharp_array_operations/array_index_of_returns_first_matching_position
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] a = {"a","b","c","b"};
__Check((System.Array.IndexOf(a,"b")).ToString(), "1");
