// vybe-test: csharp/csharp_record_types/nominal_record_with_init_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

using static __Harness;

var c = new Config { Host="localhost", Port=8080 }
;
__P((c.Host).ToString());
__P((c.Port).ToString());
__Check("localhost\n8080");

record Config { public string Host { get; init; } public int Port { get; init; } }

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
