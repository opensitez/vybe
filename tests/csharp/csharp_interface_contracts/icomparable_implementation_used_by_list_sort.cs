// vybe-test: csharp/csharp_interface_contracts/icomparable_implementation_used_by_list_sort
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

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

class Priority : System.IComparable<Priority> {
    public int Level;
    public int CompareTo(Priority other) => Level.CompareTo(other.Level);
}
var list = new System.Collections.Generic.List<Priority> {
    new Priority{Level=3}, new Priority{Level=1}, new Priority{Level=2}
};
list.Sort();
foreach(var p in list) __P((p.Level).ToString());
__Check("1\n2\n3");
