// vybe-test: csharp/csharp_properties_accessors/getter_only_computed_property_reads_backing_fields
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

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

class Rectangle {
    public int Width { get; set; }
    public int Height { get; set; }
    public int Area { get { return Width * Height; } }
}
var rectangle = new Rectangle { Width = 4, Height = 6 };
__P((rectangle.Area).ToString());
__Check("24");
