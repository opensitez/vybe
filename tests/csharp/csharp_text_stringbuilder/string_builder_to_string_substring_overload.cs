// vybe-test: csharp/csharp_text_stringbuilder/string_builder_to_string_substring_overload
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("hello world");
__Check((sb.ToString(6,5)).ToString(), "world");
