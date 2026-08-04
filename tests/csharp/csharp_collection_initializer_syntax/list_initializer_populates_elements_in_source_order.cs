// vybe-test: csharp/csharp_collection_initializer_syntax/list_initializer_populates_elements_in_source_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

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
var items = new List<int> { 3, 1, 4 };
__P((items[0]).ToString());
__P((items[2]).ToString());
__Check("3\n4");
