// vybe-test: csharp/csharp_nested_classes/nested_class_inherits_from_another_nested_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

using static __Harness;

Shapes.Base b=new Shapes.Circle();
__P((b.Name()).ToString());
__Check("circle");

class Shapes{
    public class Base{public virtual string Name()=>"shape";}
    public class Circle:Base{public override string Name()=>"circle";}
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
