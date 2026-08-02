// vybe-test: csharp/csharp_nested_classes/nested_class_inherits_from_another_nested_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shapes{
    public class Base{public virtual string Name()=>"shape";}
    public class Circle:Base{public override string Name()=>"circle";}
}
Shapes.Base b=new Shapes.Circle();
__Check((b.Name()).ToString(), "circle");
