// vybe-test: csharp/csharp_linq_deferred_execution/linq_all_short_circuits_on_first_failing_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
int probes = 0;
bool ok = new[] { 2, 4, 5, 8 }.All(x => { probes++; return x % 2 == 0; });
__Check((ok).ToString(), "False");
__Check((probes).ToString(), "3");
