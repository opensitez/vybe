// vybe-test: csharp/csharp_using_static/using_static_string_allows_unqualified_join
// origin: languages/csharp/tests/csharp/test_csharp_using_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using static System.String;
__Check((Join("-",new[]{"a","b","c"})).ToString(), "a-b-c");
