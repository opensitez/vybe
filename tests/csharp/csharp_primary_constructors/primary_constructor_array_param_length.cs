// vybe-test: csharp/csharp_primary_constructors/primary_constructor_array_param_length
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pack(int[] data) { public int Count => data.Length; }
__Check((new Pack(new[] { 1, 2, 3 }).Count).ToString(), "3");
