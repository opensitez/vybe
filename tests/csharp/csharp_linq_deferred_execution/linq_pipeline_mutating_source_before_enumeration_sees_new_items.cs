// vybe-test: csharp/csharp_linq_deferred_execution/linq_pipeline_mutating_source_before_enumeration_sees_new_items
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

using System.Collections.Generic;
using System.Linq;
var data = new List<int> { 1, 2 };
var query = data.Select(x => x * 10);
data.Add(3);
foreach (var value in query) __P((value).ToString());
__Check("10\n20\n30");
