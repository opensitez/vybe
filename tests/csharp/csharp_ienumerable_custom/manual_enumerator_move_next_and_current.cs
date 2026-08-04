// vybe-test: csharp/csharp_ienumerable_custom/manual_enumerator_move_next_and_current
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

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

var list=new System.Collections.Generic.List<int>{10,20,30};
using var e=list.GetEnumerator();
e.MoveNext();
__P((e.Current).ToString());
__Check("10");
