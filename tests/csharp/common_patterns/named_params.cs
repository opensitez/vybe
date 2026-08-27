// vybe-test: csharp/common_patterns/named_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Rect.Area(width: 5, height: 3)).ToString());
__P((Rect.Area(height: 10, width: 2)).ToString());
__Check("15\n20");

class Rect {
    public static int Area(int width, int height) { return width * height; }
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
