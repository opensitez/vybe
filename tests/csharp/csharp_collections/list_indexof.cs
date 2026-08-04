// vybe-test: csharp/csharp_collections/list_indexof
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

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
var list = new List<string>();
list.Add("a");
list.Add("b");
list.Add("c");
__P((list.IndexOf("b")).ToString());
__P((list.IndexOf("z")).ToString());
__Check("1\n-1");
