// vybe-test: csharp/csharp_events/custom_event_args_carries_data_to_handler
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

using static __Harness;

int received = 0;
var src = new Source();
src.Changed += (s, e) => received = e.Value;
src.Change(77);
__P((received).ToString());
__Check("77");

class DataArgs : System.EventArgs { public int Value; }

class Source {
    public event System.EventHandler<DataArgs> Changed;
    public void Change(int v) => Changed?.Invoke(this, new DataArgs{Value=v});
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
