// vybe-test: csharp/csharp_oop_user_defined_explicit_implicit_conversions/custom_conversions_case_14

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

CustomInt_14 c = 14;
int unboxed = (int)c;
__P(unboxed.ToString());
__Check("14");

class CustomInt_14 {
    public int Value { get; }
    public CustomInt_14(int v) => Value = v;
    public static implicit operator CustomInt_14(int v) => new CustomInt_14(v);
    public static explicit operator int(CustomInt_14 c) => c.Value;
}
