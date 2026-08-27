// vybe-test: csharp/csharp_class_indexers/indexer_parameter_can_be_enum_typed_key
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

__P((new Point()[Axis.Y]).ToString());
__Check("8");

enum Axis { X, Y }

class Point {
  int[] values = { 4, 8 };
  public int this[Axis axis] { get { return values[(int)axis]; } }
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
