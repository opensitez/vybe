// vybe-test: csharp/csharp_generics/generic_dictionary_usage
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dict = new Dictionary<string, int>();
dict.Add("x", 10);
dict.Add("y", 20);
__Check((dict["x"]).ToString(), "10");
__Check((dict.Count).ToString(), "2");
