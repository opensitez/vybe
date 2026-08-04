// vybe-test: csharp/csharp_linq_deferred_execution/linq_take_short_circuits_without_visiting_entire_source
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
int visited = 0;
var query = Enumerable.Range(1, 100).Select(x => { visited++; return x; }).Take(2);
foreach (var value in query) __P((value).ToString());
__P((visited).ToString());
__Check("1\n2\n2");
