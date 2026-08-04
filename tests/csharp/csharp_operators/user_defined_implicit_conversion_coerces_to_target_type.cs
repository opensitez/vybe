// vybe-test: csharp/csharp_operators/user_defined_implicit_conversion_coerces_to_target_type
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

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

struct Inch {
    public double Value;
    public static implicit operator double(Inch i) => i.Value;
}
double length = new Inch { Value = 2.5 };
__P((length).ToString());
__Check("2.5");
