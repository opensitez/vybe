// vybe-test: csharp/csharp_events/multiple_subscribers_all_called_in_order
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

using static __Harness;

string log = "";
var e = new Emitter();
e.Signal += v => log += "A";
e.Signal += v => log += "B";
e.Emit("x");
__P((log).ToString());
__Check("AB");

class Emitter {
    public event System.Action<string> Signal;
    public void Emit(string v) => Signal?.Invoke(v);
}

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
