// vybe-test: csharp/csharp_pattern_switch_advanced/nested_tuple_pattern_matches_pair_of_conditions
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

using static __Harness;

string Combo(bool a,bool b)=>(a,b) switch{
    (true,true)=>"both",
    (true,false)=>"left",
    (false,true)=>"right",
    _=>"none"}
;
__P((Combo(true,false)).ToString());
__Check("left");

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
