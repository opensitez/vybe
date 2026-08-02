// vybe-test: csharp/csharp_string_advanced_ops/string_format_right_align_pad_with_width
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0,10}","hello")).ToString(), "     hello");
