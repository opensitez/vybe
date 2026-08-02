// vybe-test: csharp/csharp_generic_methods/generic_method_returns_default_for_empty_sequence
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T FirstOrDefault<T>(T[] arr)=>arr.Length>0?arr[0]:default;
__Check((FirstOrDefault(new int[]{})).ToString(), "0");
__Check((FirstOrDefault(new[]{9})).ToString(), "9");
