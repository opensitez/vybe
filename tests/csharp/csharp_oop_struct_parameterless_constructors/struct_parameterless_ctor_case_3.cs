// vybe-test: csharp/csharp_oop_struct_parameterless_constructors/struct_parameterless_ctor_case_3

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

var p = new PointStruct_3();
__P(p.X.ToString());
__Check("3");

struct PointStruct_3 {
    public int X { get; } = 3;
    public PointStruct_3() { }
}
