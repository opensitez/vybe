// vybe-test: csharp/strings_advanced/string_tochararray
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] chars = "hello".ToCharArray();
Array.Reverse(chars);
__Check((new string(chars)).ToString(), "olleh");
