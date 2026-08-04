// vybe-test: csharp/csharp_delegates_advanced/delegate_stored_in_dictionary_and_dispatched_by_key
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

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

var ops=new System.Collections.Generic.Dictionary<string,System.Func<int,int,int>>{
    {"add",(a,b)=>a+b},
    {"mul",(a,b)=>a*b}
};
__P((ops["add"](3,4)).ToString());
__P((ops["mul"](3,4)).ToString());
__Check("7\n12");
