// vybe-test: csharp/csharp_custom_event_accessors/custom_event_action_int
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

int got=0;
var s=new Src();
s.Value+=v=>got=v;
s.Set(15);
__P((got).ToString());
__Check("15");

class Src{System.Action<int> _h; public event System.Action<int> Value{add{_h+=value;} remove{_h-=value;}} public void Set(int v){_h?.Invoke(v);}}

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
