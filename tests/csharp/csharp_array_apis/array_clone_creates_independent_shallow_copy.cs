// vybe-test: csharp/csharp_array_apis/array_clone_creates_independent_shallow_copy
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var source = new[] { 1, 2 }; var clone = (int[])source.Clone(); clone[0] = 9; __Check((source[0]).ToString(), "1"); __Check((clone[0]).ToString(), "9");
