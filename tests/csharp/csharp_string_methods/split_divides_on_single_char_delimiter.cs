// vybe-test: csharp/csharp_string_methods/split_divides_on_single_char_delimiter
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p = "a,b,c".Split(','); __Check((p[1]).ToString(), "b");
