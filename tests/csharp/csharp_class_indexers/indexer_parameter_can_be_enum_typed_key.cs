// vybe-test: csharp/csharp_class_indexers/indexer_parameter_can_be_enum_typed_key
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

enum Axis { X, Y }
class Point {
  int[] values = { 4, 8 };
  public int this[Axis axis] { get { return values[(int)axis]; } }
}
__P((new Point()[Axis.Y]).ToString());
__Check("8");
