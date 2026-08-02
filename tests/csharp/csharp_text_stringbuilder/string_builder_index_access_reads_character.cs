// vybe-test: csharp/csharp_text_stringbuilder/string_builder_index_access_reads_character
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder("hello");
__Check((sb[1]).ToString(), "e");
