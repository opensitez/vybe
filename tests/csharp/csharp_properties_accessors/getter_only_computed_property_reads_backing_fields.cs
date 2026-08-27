// vybe-test: csharp/csharp_properties_accessors/getter_only_computed_property_reads_backing_fields
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var rectangle = new Rectangle { Width = 4, Height = 6 }
;
__P((rectangle.Area).ToString());
__Check("24");

class Rectangle {
    public int Width { get; set; }
    public int Height { get; set; }
    public int Area { get { return Width * Height; } }
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
