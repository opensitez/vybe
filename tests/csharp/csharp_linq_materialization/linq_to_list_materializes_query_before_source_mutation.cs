// vybe-test: csharp/csharp_linq_materialization/linq_to_list_materializes_query_before_source_mutation
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

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
var source = new List<int> { 1, 2 };
var snapshot = source.Select(x => x).ToList();
source.Add(3);
__P((snapshot.Count).ToString());
__P((source.Count).ToString());
__Check("2\n3");
