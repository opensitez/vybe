// vybe-test: csharp/csharp_attributes/obsolete_attribute_is_standard_bcl_attribute
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

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

class Old{
    [System.Obsolete("use NewMethod")]
    public void OldMethod(){}
}
var mi=typeof(Old).GetMethod("OldMethod");
bool hasObs=mi.GetCustomAttributes(typeof(System.ObsoleteAttribute),false).Length>0;
__P((hasObs).ToString());
__Check("True");
