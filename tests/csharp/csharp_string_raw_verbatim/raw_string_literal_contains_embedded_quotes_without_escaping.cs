// vybe-test: csharp/csharp_string_raw_verbatim/raw_string_literal_contains_embedded_quotes_without_escaping
// origin: languages/csharp/tests/csharp/test_csharp_string_raw_verbatim.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s="""She said "hello" to him.""";
__Check((s.Contains(""hello"")).ToString(), "True");
