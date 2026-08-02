// vybe-test: csharp/csharp_string_advanced_ops/string_format_left_align_with_negative_width
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0,-10}|","hello")).ToString(), "hello     |");
