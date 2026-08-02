// vybe-test: csharp/advanced/array_length
// origin: languages/csharp/tests/csharp/test_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = new int[] { 1, 2, 3 };
        __Check((arr[0]).ToString(), "1");
        __Check((arr[2]).ToString(), "3");
