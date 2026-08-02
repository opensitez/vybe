// vybe-test: csharp/csharp_string_methods/pad_right_left_aligns_to_minimum_width
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hi".PadRight(5) + "|").ToString(), "hi   |");
