// vybe-test: csharp/csharp_custom_event_accessors/custom_event_subscribe_after_first_raise
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

var b=new Btn();
b.Raise();
b.Click+=()=>__P(("late").ToString());
b.Raise();
__Check("late");

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}

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
