// vybe-test: csharp/advanced/array_creation_and_index
// origin: languages/csharp/tests/csharp/test_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = new int[] { 10, 20, 30 };
        __Check((arr[1]).ToString(), "20");
