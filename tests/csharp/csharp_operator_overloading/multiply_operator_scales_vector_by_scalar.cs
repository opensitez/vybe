// vybe-test: csharp/csharp_operator_overloading/multiply_operator_scales_vector_by_scalar
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

struct Vec{public int X;
public static Vec operator*(Vec v,int s)=>new Vec{X=v.X*s};}
__P(((new Vec{X=3}*4).X).ToString());
__Check("12");
