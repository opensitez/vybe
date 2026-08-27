// vybe-test: csharp/csharp_enum_operations/enum_parse_converts_string_name_to_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

using static __Harness;

var d = (Day)System.Enum.Parse(typeof(Day),"Wed");
__P((d).ToString());
__Check("Wed");

enum Day{Mon,Tue,Wed,Thu,Fri}

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
