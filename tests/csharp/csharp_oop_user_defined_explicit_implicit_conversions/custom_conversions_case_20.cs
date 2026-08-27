// vybe-test: csharp/csharp_oop_user_defined_explicit_implicit_conversions/custom_conversions_case_20

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

CustomInt_20 c = 20;
int unboxed = (int)c;
__P(unboxed.ToString());
__Check("20");

class CustomInt_20 {
    public int Value { get; }
    public CustomInt_20(int v) => Value = v;
    public static implicit operator CustomInt_20(int v) => new CustomInt_20(v);
    public static explicit operator int(CustomInt_20 c) => c.Value;
}
