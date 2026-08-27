// vybe-test: csharp/csharp_event_advanced/unsubscribing_from_event_stops_handler_firing
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

using static __Harness;

__P("Valid_unsubscribing_from_event_stops_handler_firing");
__Check("Valid_unsubscribing_from_event_stops_handler_firing");
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
