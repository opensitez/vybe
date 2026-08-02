// vybe-test: csharp/csharp_arrays_advanced/new_array_syntax
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = new[] { 10, 20, 30, 40, 50 };
__Check((arr.Length).ToString(), "5");
__Check((arr[2]).ToString(), "30");
