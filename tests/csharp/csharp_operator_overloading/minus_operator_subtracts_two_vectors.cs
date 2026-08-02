// vybe-test: csharp/csharp_operator_overloading/minus_operator_subtracts_two_vectors
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec{public int X,Y;
public static Vec operator-(Vec a,Vec b)=>new Vec{X=a.X-b.X,Y=a.Y-b.Y};}
var v=new Vec{X=5,Y=3}-new Vec{X=2,Y=1};
__Check((v.X).ToString(), "3");
