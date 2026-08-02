// vybe-test: csharp/csharp_properties_accessors/getter_only_computed_property_reads_backing_fields
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Rectangle {
    public int Width { get; set; }
    public int Height { get; set; }
    public int Area { get { return Width * Height; } }
}
var rectangle = new Rectangle { Width = 4, Height = 6 };
__Check((rectangle.Area).ToString(), "24");
