// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_inherits_nested_base
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shapes{public class Base{public virtual string Name()=>"base";} public class Circle:Base{public override string Name()=>"circle";}} __Check((new Shapes.Circle().Name()).ToString(), "circle");
