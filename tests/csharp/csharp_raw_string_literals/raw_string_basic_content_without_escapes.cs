// vybe-test: csharp/csharp_raw_string_literals/raw_string_basic_content_without_escapes
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""hello"""; __Check((text).ToString(), "hello");
