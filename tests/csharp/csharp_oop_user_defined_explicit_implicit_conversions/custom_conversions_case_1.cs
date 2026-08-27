// vybe-test: csharp/csharp_oop_user_defined_explicit_implicit_conversions/custom_conversions_case_1

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

CustomInt_1 c = 1;
int unboxed = (int)c;
__P(unboxed.ToString());
__Check("1");

class CustomInt_1 {
    public int Value { get; }
    public CustomInt_1(int v) => Value = v;
    public static implicit operator CustomInt_1(int v) => new CustomInt_1(v);
    public static explicit operator int(CustomInt_1 c) => c.Value;
}
