// vybe-test: csharp/csharp_string_ops_advanced/split_with_max_count_limits_resulting_segments
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts="a:b:c:d".Split(':',2);
__Check((parts.Length).ToString(), "2"); __Check((parts[1]).ToString(), "b:c:d");
