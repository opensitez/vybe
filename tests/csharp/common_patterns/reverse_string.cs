// vybe-test: csharp/common_patterns/reverse_string
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "Hello World";
char[] chars = s.ToCharArray();
Array.Reverse(chars);
__Check((new string(chars)).ToString(), "dlroW olleH");
