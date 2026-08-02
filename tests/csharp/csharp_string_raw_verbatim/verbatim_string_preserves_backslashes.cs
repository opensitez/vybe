// vybe-test: csharp/csharp_string_raw_verbatim/verbatim_string_preserves_backslashes
// origin: languages/csharp/tests/csharp/test_csharp_string_raw_verbatim.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path=@"C:\Users\test\file.txt";
__Check((path.Contains(@"\test")).ToString(), "True");
