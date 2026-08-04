// vybe-test: csharp/csharp_equality_contracts/list_reference_equality_is_false_for_distinct_instances_with_same_contents
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

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
var left = new List<int> { 1, 2 };
var right = new List<int> { 1, 2 };
__P((left == right).ToString());
__P((left.SequenceEqual(right)).ToString());
__Check("False\nTrue");
