// vybe-test: csharp/csharp_linq_deferred_execution/linq_zip_pairs_elements_until_shorter_sequence_ends
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
foreach (var pair in new[] { 1, 2, 3 }.Zip(new[] { 10, 20 }, (a, b) => a + b)) __P((pair).ToString());
__Check("11\n22");
