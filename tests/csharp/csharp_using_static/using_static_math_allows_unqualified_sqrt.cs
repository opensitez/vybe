// vybe-test: csharp/csharp_using_static/using_static_math_allows_unqualified_sqrt
// origin: languages/csharp/tests/csharp/test_csharp_using_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using static System.Math;
__Check((Sqrt(16)).ToString(), "4");
