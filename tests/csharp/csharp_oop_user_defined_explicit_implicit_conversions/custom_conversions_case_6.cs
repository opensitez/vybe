// vybe-test: csharp/csharp_oop_user_defined_explicit_implicit_conversions/custom_conversions_case_6

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

CustomInt_6 c = 6;
int unboxed = (int)c;
__P(unboxed.ToString());
__Check("6");

class CustomInt_6 {
    public int Value { get; }
    public CustomInt_6(int v) => Value = v;
    public static implicit operator CustomInt_6(int v) => new CustomInt_6(v);
    public static explicit operator int(CustomInt_6 c) => c.Value;
}
