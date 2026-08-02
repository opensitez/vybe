// vybe-test: csharp/strings_advanced/verbatim_string
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path = @"C:\Users\test\file.txt";
__Check((path).ToString(), "C:\\Users\\test\\file.txt");
