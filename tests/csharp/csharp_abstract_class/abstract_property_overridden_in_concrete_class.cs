// vybe-test: csharp/csharp_abstract_class/abstract_property_overridden_in_concrete_class
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Shape{public abstract double Area;}
class Square:Shape{public double Side;public override double Area=>Side*Side;}
Shape s=new Square{Side=4};
__Check((s.Area).ToString(), "16");
