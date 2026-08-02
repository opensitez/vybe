// vybe-test: csharp/csharp_linq_materialization/linq_single_requires_exactly_one_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
__Check((new[] { 42 }.Single()).ToString(), "42");
