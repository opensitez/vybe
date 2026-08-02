// vybe-test: csharp/csharp_linq_deferred_execution/linq_any_short_circuits_on_first_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
int probes = 0;
bool found = new[] { 1, 2, 3 }.Any(x => { probes++; return x == 2; });
__Check((found).ToString(), "True");
__Check((probes).ToString(), "2");
