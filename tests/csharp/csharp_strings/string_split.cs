// vybe-test: csharp/csharp_strings/string_split
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] parts = "a,b,c".Split(",");
__Check((parts.Length).ToString(), "3");
__Check((parts[1]).ToString(), "b");
