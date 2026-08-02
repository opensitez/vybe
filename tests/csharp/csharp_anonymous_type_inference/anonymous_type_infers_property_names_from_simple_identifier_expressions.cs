// vybe-test: csharp/csharp_anonymous_type_inference/anonymous_type_infers_property_names_from_simple_identifier_expressions
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_type_inference.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int width = 4;
string label = "box";
var shape = new { width, label };
__Check((shape.width).ToString(), "4");
__Check((shape.label).ToString(), "box");
