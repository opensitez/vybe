// vybe-test: csharp/csharp_string_format/string_format_with_negative_alignment_left_pads_to_field_width
// origin: languages/csharp/tests/csharp/test_csharp_string_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0,-5}|", "ab")).ToString(), "ab   |");
