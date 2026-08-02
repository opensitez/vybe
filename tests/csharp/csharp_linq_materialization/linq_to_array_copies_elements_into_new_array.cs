// vybe-test: csharp/csharp_linq_materialization/linq_to_array_copies_elements_into_new_array
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
var copy = new[] { 5, 6 }.Select(x => x).ToArray();
__Check((copy[1]).ToString(), "6");
