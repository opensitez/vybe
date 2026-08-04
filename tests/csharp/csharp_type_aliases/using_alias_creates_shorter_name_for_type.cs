// vybe-test: csharp/csharp_type_aliases/using_alias_creates_shorter_name_for_type
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

using IntList=System.Collections.Generic.List<int>;
var list=new IntList{1,2,3};
__P((list.Count).ToString());
__Check("3");
