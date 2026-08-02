// vybe-test: csharp/csharp_strings_ext/verbatim_string
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var path = @"C:\Users\test\file.txt";
__Check((path).ToString(), "C:\\Users\\test\\file.txt");
