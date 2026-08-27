// vybe-test: csharp/csharp_custom_event_accessors/custom_event_eventhandler_generic
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

string out_="";
var c=new Ch();
c.Sent+=(o,e)=>out_=e.Text;
c.Emit("hi");
__P((out_).ToString());
__Check("hi");

class Msg: System.EventArgs{public string Text;}

class Ch{System.EventHandler<Msg> _h; public event System.EventHandler<Msg> Sent{add{_h+=value;} remove{_h-=value;}} public void Emit(string t){_h?.Invoke(this,new Msg{Text=t});}}

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
