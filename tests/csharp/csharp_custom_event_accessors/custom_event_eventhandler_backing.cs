// vybe-test: csharp/csharp_custom_event_accessors/custom_event_eventhandler_backing
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

int n=0;
var s=new Src();
s.Changed+=(o,e)=>n++;
s.Raise();
__P((n).ToString());
__Check("1");

class Src{System.EventHandler _h; public event System.EventHandler Changed{add{_h+=value;} remove{_h-=value;}} public void Raise(){_h?.Invoke(this,System.EventArgs.Empty);}}

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
