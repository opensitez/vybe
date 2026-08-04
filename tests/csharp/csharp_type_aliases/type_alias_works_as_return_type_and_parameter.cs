// vybe-test: csharp/csharp_type_aliases/type_alias_works_as_return_type_and_parameter
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

using NameMap=System.Collections.Generic.Dictionary<string,string>;
NameMap Build()=>new NameMap{{"k","v"}};
__P((Build()["k"]).ToString());
__Check("v");
