// vybe-test: csharp/common_patterns/enum_in_switch
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

Season s = Season.Summer;
switch (s) {
    case Season.Spring: __P(("spring").ToString()); break;
    case Season.Summer: __P(("summer").ToString()); break;
    case Season.Autumn: __P(("autumn").ToString()); break;
    case Season.Winter: __P(("winter").ToString()); break;
}
__Check("summer");

enum Season { Spring, Summer, Autumn, Winter }

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
