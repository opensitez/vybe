// vybe-test: csharp/csharp_oop_user_defined_explicit_implicit_conversions/custom_conversions_case_15

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

CustomInt_15 c = 15;
int unboxed = (int)c;
__P(unboxed.ToString());
__Check("15");

class CustomInt_15 {
    public int Value { get; }
    public CustomInt_15(int v) => Value = v;
    public static implicit operator CustomInt_15(int v) => new CustomInt_15(v);
    public static explicit operator int(CustomInt_15 c) => c.Value;
}
