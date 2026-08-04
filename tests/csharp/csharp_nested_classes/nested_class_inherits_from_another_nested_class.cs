// vybe-test: csharp/csharp_nested_classes/nested_class_inherits_from_another_nested_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

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

class Shapes{
    public class Base{public virtual string Name()=>"shape";}
    public class Circle:Base{public override string Name()=>"circle";}
}
Shapes.Base b=new Shapes.Circle();
__P((b.Name()).ToString());
__Check("circle");
