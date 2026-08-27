// vybe-test: csharp/csharp_oop_struct_parameterless_constructors/struct_parameterless_ctor_case_17

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

var p = new PointStruct_17();
__P(p.X.ToString());
__Check("17");

struct PointStruct_17 {
    public int X { get; } = 17;
    public PointStruct_17() { }
}
