// vybe-test: csharp/csharp_custom_event_accessors/custom_event_base_backing
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

int n=0;
var c=new Child();
c.Ping+=()=>n++;
c.Fire();
__P((n).ToString());
__Check("1");

class Base{System.Action _e; public event System.Action Ping{add{_e+=value;} remove{_e-=value;}} protected void OnPing(){_e?.Invoke();}}

class Child:Base{public void Fire(){OnPing();}}

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
