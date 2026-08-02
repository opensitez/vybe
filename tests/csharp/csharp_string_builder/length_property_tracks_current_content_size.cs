// vybe-test: csharp/csharp_string_builder/length_property_tracks_current_content_size
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("abc");
__Check((sb.Length).ToString(), "3");
