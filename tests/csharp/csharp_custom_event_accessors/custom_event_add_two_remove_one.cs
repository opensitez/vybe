// vybe-test: csharp/csharp_custom_event_accessors/custom_event_add_two_remove_one
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

using static __Harness;

System.Action a=()=>{}
;
System.Action b=()=>{}
;
var btn=new Btn();
btn.Click+=a;
btn.Click+=b;
btn.Click-=a;
__P((btn.Count).ToString());
__Check("1");

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public int Count=>_c==null?0:_c.GetInvocationList().Length;}

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
