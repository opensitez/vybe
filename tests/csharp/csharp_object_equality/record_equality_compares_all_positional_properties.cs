// vybe-test: csharp/csharp_object_equality/record_equality_compares_all_positional_properties
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y);
var a = new Point(1, 2);
var b = new Point(1, 2);
var c = new Point(1, 3);
__Check((a.Equals(b)).ToString(), "True");
__Check((a.Equals(c)).ToString(), "False");
