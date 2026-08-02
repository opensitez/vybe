// vybe-test: csharp/csharp_string_builder/clear_resets_length_to_zero
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("data");
sb.Clear();
__Check((sb.Length).ToString(), "0");
