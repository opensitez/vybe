// vybe-test: csharp/csharp_type_aliases/using_alias_for_fully_qualified_type
// origin: languages/csharp/tests/csharp/test_csharp_type_aliases.rs

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

using Dict=System.Collections.Generic.Dictionary<string,int>;
var d=new Dict{{"a",1},{"b",2}};
__P((d["b"]).ToString());
__Check("2");
