// vybe-test: csharp/csharp_oop_ref_struct_interfaces/ref_struct_interface_case_4

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

var holder = new RefSpanHolder_4 { Span = new int[] { 4 } };
__P(holder.Span[0].ToString());
holder.Dispose();
__Check("4");

ref struct RefSpanHolder_4 : IDisposable {
    public ReadOnlySpan<int> Span;
    public void Dispose() { }
}
