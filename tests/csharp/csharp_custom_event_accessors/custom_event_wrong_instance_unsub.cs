// vybe-test: csharp/csharp_custom_event_accessors/custom_event_wrong_instance_unsub
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

int n=0;
System.Action h=()=>n++;
var a=new Btn();
var b=new Btn();
a.Click+=h;
b.Click-=h;
a.Raise();
__P((n).ToString());
__Check("1");

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
