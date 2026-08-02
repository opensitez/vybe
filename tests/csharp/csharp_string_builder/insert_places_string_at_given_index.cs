// vybe-test: csharp/csharp_string_builder/insert_places_string_at_given_index
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("ac");
sb.Insert(1,"b");
__Check((sb.ToString()).ToString(), "abc");
