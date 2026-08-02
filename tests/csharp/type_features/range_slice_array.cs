// vybe-test: csharp/type_features/range_slice_array
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = new int[] { 10, 20, 30, 40, 50 };
        var sub = arr[1..3];
        __Check((sub[0]).ToString(), "20");
        __Check((sub[1]).ToString(), "30");
