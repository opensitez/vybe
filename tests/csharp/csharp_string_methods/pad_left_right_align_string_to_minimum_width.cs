// vybe-test: csharp/csharp_string_methods/pad_left_right_align_string_to_minimum_width
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hi".PadLeft(5)).ToString(), "   hi");
