// vybe-test: csharp/csharp_deferred_execution/to_list_materialises_and_snapshot_captures_source
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

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

var source=new System.Collections.Generic.List<int>{1,2,3};
var snapshot=source.ToList();
source.Add(4);
__P((snapshot.Count).ToString());
__Check("3");
