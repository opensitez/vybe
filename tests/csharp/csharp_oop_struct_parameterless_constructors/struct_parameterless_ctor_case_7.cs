// vybe-test: csharp/csharp_oop_struct_parameterless_constructors/struct_parameterless_ctor_case_7

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

var p = new PointStruct_7();
__P(p.X.ToString());
__Check("7");

struct PointStruct_7 {
    public int X { get; } = 7;
    public PointStruct_7() { }
}
