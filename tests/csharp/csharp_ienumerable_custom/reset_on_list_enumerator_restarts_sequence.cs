// vybe-test: csharp/csharp_ienumerable_custom/reset_on_list_enumerator_restarts_sequence
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

var list=new System.Collections.Generic.List<int>{1,2,3};
int count=0;
foreach(var _ in list) count++;
foreach(var _ in list) count++;
__P((count).ToString());
__Check("6");
