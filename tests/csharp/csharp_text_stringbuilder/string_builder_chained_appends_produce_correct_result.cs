// vybe-test: csharp/csharp_text_stringbuilder/string_builder_chained_appends_produce_correct_result
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb=new System.Text.StringBuilder();
sb.Append("a").Append("b").Append("c");
__Check((sb.ToString()).ToString(), "abc");
