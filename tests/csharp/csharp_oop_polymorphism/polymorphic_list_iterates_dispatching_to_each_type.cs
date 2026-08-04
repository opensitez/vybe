// vybe-test: csharp/csharp_oop_polymorphism/polymorphic_list_iterates_dispatching_to_each_type
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

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

abstract class Shape{public abstract int Size();}
class Square:Shape{public override int Size()=>4;}
class Triangle:Shape{public override int Size()=>3;}
var shapes=new System.Collections.Generic.List<Shape>{new Square(),new Triangle(),new Square()};
int sum=0; foreach(var s in shapes) sum+=s.Size();
__P((sum).ToString());
__Check("11");
