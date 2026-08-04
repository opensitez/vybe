// vybe-test: csharp/csharp_operator_overloading/explicit_conversion_to_primitive_requires_cast
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

struct Percent{public double Value;
public static explicit operator double(Percent p)=>p.Value/100.0;}
var p=new Percent{Value=50};
__P(((double)p).ToString());
__Check("0.5");
