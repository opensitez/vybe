// vybe-test: csharp/csharp_string_builder/replace_substitutes_all_occurrences_in_buffer
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("aabbaa");
sb.Replace("aa","X");
__Check((sb.ToString()).ToString(), "XbbX");
