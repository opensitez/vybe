// vybe-test: csharp/csharp_class_features/memberwise_clone_creates_shallow_copy
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point:System.ICloneable{public int X,Y;public object Clone()=>MemberwiseClone();}
var a=new Point{X=1,Y=2};
var b=(Point)a.Clone();
b.X=99;
__Check((a.X).ToString(), "1");
