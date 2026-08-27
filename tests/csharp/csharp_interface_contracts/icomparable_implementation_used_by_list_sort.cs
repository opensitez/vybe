// vybe-test: csharp/csharp_interface_contracts/icomparable_implementation_used_by_list_sort
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

using static __Harness;

var list = new System.Collections.Generic.List<Priority> {
    new Priority{Level=3}, new Priority{Level=1}, new Priority{Level=2}
}
;
list.Sort();
foreach(var p in list) __P((p.Level).ToString());
__Check("1\n2\n3");

class Priority : System.IComparable<Priority> {
    public int Level;
    public int CompareTo(Priority other) => Level.CompareTo(other.Level);
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
