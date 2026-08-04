// vybe-test: csharp/csharp_access_modifiers/private_setter_means_field_read_only_from_outside
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

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

class Counter{
    public int Count{get;private set;}
    public void Tick(){Count++;}
}
var c=new Counter(); c.Tick(); c.Tick();
__P((c.Count).ToString());
__Check("2");
