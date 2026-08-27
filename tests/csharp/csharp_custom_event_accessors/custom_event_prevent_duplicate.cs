// vybe-test: csharp/csharp_custom_event_accessors/custom_event_prevent_duplicate
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

int n=0;
System.Action h=()=>n++;
var b=new Btn();
b.Click+=h;
b.Click+=h;
b.Raise();
__P((n).ToString());
__Check("1");

class Btn{System.Action _c; public event System.Action Click{add{if(_c==null||!_c.GetInvocationList().Contains(value)) _c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}

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
