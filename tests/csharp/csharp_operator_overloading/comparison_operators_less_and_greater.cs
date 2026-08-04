// vybe-test: csharp/csharp_operator_overloading/comparison_operators_less_and_greater
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

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

class Weight:System.IComparable<Weight>{
public int Kg;
public static bool operator<(Weight a,Weight b)=>a.Kg<b.Kg;
public static bool operator>(Weight a,Weight b)=>a.Kg>b.Kg;
public int CompareTo(Weight o)=>Kg.CompareTo(o.Kg);}
var a=new Weight{Kg=5}; var b=new Weight{Kg=10};
__P((a<b).ToString()); __P((a>b).ToString());
__Check("True\nFalse");
