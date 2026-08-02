// vybe-test: csharp/csharp_string_builder/remove_deletes_character_range_by_start_and_count
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("hello");
sb.Remove(1,3);
__Check((sb.ToString()).ToString(), "ho");
