// vybe-test: csharp/csharp_expression_bodied/expression_bodied_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle{public double R;public double Area=>System.Math.PI*R*R;}
__Check((System.Math.Round(new Circle{R=0}.Area)).ToString(), "0");
