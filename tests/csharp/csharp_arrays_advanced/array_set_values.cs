// vybe-test: csharp/csharp_arrays_advanced/array_set_values
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = new int[3];
arr[0] = 10;
arr[1] = 20;
arr[2] = 30;
__Check((arr[0] + arr[1] + arr[2]).ToString(), "60");
