// vybe-test: csharp/csharp_operator_overloading/comparison_operators_less_and_greater
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Weight:System.IComparable<Weight>{
public int Kg;
public static bool operator<(Weight a,Weight b)=>a.Kg<b.Kg;
public static bool operator>(Weight a,Weight b)=>a.Kg>b.Kg;
public int CompareTo(Weight o)=>Kg.CompareTo(o.Kg);}
var a=new Weight{Kg=5}; var b=new Weight{Kg=10};
__Check((a<b).ToString(), "True"); __Check((a>b).ToString(), "False");
