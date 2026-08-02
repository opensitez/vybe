// vybe-test: csharp/csharp_strings_ext/string_builder_insert_puts_text_at_offset
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("ac");
sb.Insert(1, "b");
__Check((sb.ToString()).ToString(), "abc");
