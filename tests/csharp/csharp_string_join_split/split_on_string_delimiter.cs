// vybe-test: csharp/csharp_string_join_split/split_on_string_delimiter
// origin: languages/csharp/tests/csharp/test_csharp_string_join_split.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts="one::two::three".Split("::");
__Check((parts.Length).ToString(), "3"); __Check((parts[2]).ToString(), "three");
