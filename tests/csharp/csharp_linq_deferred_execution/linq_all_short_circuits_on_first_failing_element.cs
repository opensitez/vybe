// vybe-test: csharp/csharp_linq_deferred_execution/linq_all_short_circuits_on_first_failing_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
int probes = 0;
bool ok = new[] { 2, 4, 5, 8 }.All(x => { probes++; return x % 2 == 0; });
__P((ok).ToString());
__P((probes).ToString());
__Check("False\n3");
