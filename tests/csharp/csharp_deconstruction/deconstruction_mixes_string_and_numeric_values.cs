// vybe-test: csharp/csharp_deconstruction/deconstruction_mixes_string_and_numeric_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (name, age) = ("Grace", 42);
__Check((name + ":" + age).ToString(), "Grace:42");
