// vybe-test: csharp/csharp_dynamic/dynamic_dictionary_access_via_expando
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

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

dynamic e=new System.Dynamic.ExpandoObject();
var dict=(System.Collections.Generic.IDictionary<string,object>)e;
dict["x"]=99;
__P((e.x).ToString());
__Check("99");
