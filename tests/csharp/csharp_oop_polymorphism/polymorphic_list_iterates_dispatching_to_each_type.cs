// vybe-test: csharp/csharp_oop_polymorphism/polymorphic_list_iterates_dispatching_to_each_type
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

abstract class Shape{public abstract int Size();}
class Square:Shape{public override int Size()=>4;}
class Triangle:Shape{public override int Size()=>3;}
var shapes=new System.Collections.Generic.List<Shape>{new Square(),new Triangle(),new Square()};
int sum=0; foreach(var s in shapes) sum+=s.Size();
Console.WriteLine(sum);
