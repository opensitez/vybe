// vybe-test: csharp/strings_advanced/convert_toint32_tostring
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = Convert.ToInt32("123");
string s = Convert.ToString(456);
__Check((x).ToString(), "123");
__Check((s).ToString(), "456");
