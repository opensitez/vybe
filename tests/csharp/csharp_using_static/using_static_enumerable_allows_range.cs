// vybe-test: csharp/csharp_using_static/using_static_enumerable_allows_range
// origin: languages/csharp/tests/csharp/test_csharp_using_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using static System.Linq.Enumerable;
__Check((string.Join(",",Range(1,4))).ToString(), "1,2,3,4");
