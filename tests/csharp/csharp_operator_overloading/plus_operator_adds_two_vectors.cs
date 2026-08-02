// vybe-test: csharp/csharp_operator_overloading/plus_operator_adds_two_vectors
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec{public int X,Y;
public static Vec operator+(Vec a,Vec b)=>new Vec{X=a.X+b.X,Y=a.Y+b.Y};}
var v=new Vec{X=1,Y=2}+new Vec{X=3,Y=4};
__Check((v.X).ToString(), "4"); __Check((v.Y).ToString(), "6");
