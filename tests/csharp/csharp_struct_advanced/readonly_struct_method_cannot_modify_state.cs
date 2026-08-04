// vybe-test: csharp/csharp_struct_advanced/readonly_struct_method_cannot_modify_state
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Counter{
    public readonly int Value;
    public Counter(int v){Value=v;}
    public Counter Increment()=>new Counter(Value+1);
}
var c=new Counter(5).Increment();
__P((c.Value).ToString());
__Check("6");
