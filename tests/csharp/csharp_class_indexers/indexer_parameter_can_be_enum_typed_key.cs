// vybe-test: csharp/csharp_class_indexers/indexer_parameter_can_be_enum_typed_key
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Axis { X, Y }
class Point {
  int[] values = { 4, 8 };
  public int this[Axis axis] { get { return values[(int)axis]; } }
}
__Check((new Point()[Axis.Y]).ToString(), "8");
