// vybe-test: csharp/csharp_oop_user_defined_explicit_implicit_conversions/custom_conversions_case_16

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

CustomInt_16 c = 16;
int unboxed = (int)c;
__P(unboxed.ToString());
__Check("16");

class CustomInt_16 {
    public int Value { get; }
    public CustomInt_16(int v) => Value = v;
    public static implicit operator CustomInt_16(int v) => new CustomInt_16(v);
    public static explicit operator int(CustomInt_16 c) => c.Value;
}
