// vybe-test: csharp/csharp_operator_overloading/equality_operator_compares_value_type_fields
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Color{public int R,G,B;
public static bool operator==(Color a,Color b)=>a.R==b.R&&a.G==b.G&&a.B==b.B;
public static bool operator!=(Color a,Color b)=>!(a==b);
public override int GetHashCode()=>0; public override bool Equals(object o)=>o is Color c&&c==this;}
var a=new Color{R=1,G=2,B=3}; var b=new Color{R=1,G=2,B=3};
__Check((a==b).ToString(), "True"); __Check((a!=b).ToString(), "False");
