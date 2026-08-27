// vybe-test: csharp/csharp_record_struct_deep/record_struct_enum_with
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var j=new Job(Level.Low);
var k=j with{Tier=Level.High}
;
__P((k.Tier).ToString());
__Check("High");

enum Level{Low,High}

record struct Job(Level Tier);

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
