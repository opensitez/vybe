// vybe-test: csharp/csharp_collections/stack_operations
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
var s = new Stack<int>();
s.Push(1);
s.Push(2);
s.Push(3);
__P((s.Count).ToString());
__P((s.Pop()).ToString());
__P((s.Peek()).ToString());
__Check("3\n3\n2");
