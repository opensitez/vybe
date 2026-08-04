// vybe-test: csharp/csharp_operator_overloading/implicit_conversion_from_int_to_custom_type
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

struct Meters{public double Value;
public static implicit operator Meters(double d)=>new Meters{Value=d};}
Meters m=3.5;
__P((m.Value).ToString());
__Check("3.5");
