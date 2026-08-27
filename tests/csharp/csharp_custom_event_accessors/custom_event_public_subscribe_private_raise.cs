// vybe-test: csharp/csharp_custom_event_accessors/custom_event_public_subscribe_private_raise
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

int v=0;
var h=new Hub();
h.Signal+=()=>v=9;
h.Pulse();
__P((v).ToString());
__Check("9");

class Hub{System.Action _h; public event System.Action Signal{add{_h+=value;} remove{_h-=value;}} public void Pulse(){_h?.Invoke();}}

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
