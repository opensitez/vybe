// vybe-test: csharp/csharp_custom_event_accessors/custom_event_count_tracked
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

var b=new Btn();
System.Action h=()=>{}
;
b.Tick+=h;
b.Tick+=()=>{}
;
b.Tick-=h;
__P((b.Count).ToString());
__Check("1");

class Btn{System.Action _e; int _count; public event System.Action Tick{add{_e+=value;_count++;} remove{_e-=value;_count--;}} public int Count=>_count; public void Fire(){_e?.Invoke();}}

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
