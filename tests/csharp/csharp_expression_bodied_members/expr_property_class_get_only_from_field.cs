// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_get_only_from_field
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle { public double R = 2.0; public double Area => System.Math.PI * R * R; }
__Check((System.Math.Round(new Circle().Area, 2)).ToString(), "12.57");
