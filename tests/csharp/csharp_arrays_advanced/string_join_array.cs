// vybe-test: csharp/csharp_arrays_advanced/string_join_array
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = new[] { "a", "b", "c" };
__Check((string.Join(",", arr)).ToString(), "a,b,c");
__Check((string.Join(" - ", arr)).ToString(), "a - b - c");
