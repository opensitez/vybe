// vybe-test: csharp/csharp_string_ops_advanced/split_with_multiple_delimiters
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts="a,b;c".Split(new char[]{',',';'});
__Check((parts.Length).ToString(), "3"); __Check((parts[2]).ToString(), "c");
