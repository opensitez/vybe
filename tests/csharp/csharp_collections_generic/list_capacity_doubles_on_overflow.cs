// vybe-test: csharp/csharp_collections_generic/list_capacity_doubles_on_overflow
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

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

var list=new System.Collections.Generic.List<int>(4);
for(int i=0;i<8;i++) list.Add(i);
__P((list.Count).ToString()); __P((list.Capacity>=8).ToString());
__Check("8\nTrue");
