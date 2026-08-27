// vybe-test: csharp/csharp_oop_user_defined_explicit_implicit_conversions/custom_conversions_case_18

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

CustomInt_18 c = 18;
int unboxed = (int)c;
__P(unboxed.ToString());
__Check("18");

class CustomInt_18 {
    public int Value { get; }
    public CustomInt_18(int v) => Value = v;
    public static implicit operator CustomInt_18(int v) => new CustomInt_18(v);
    public static explicit operator int(CustomInt_18 c) => c.Value;
}
