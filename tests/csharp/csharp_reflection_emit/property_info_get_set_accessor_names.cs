// vybe-test: csharp/csharp_reflection_emit/property_info_get_set_accessor_names
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

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

class Model{public int Value{get;set;}}
var pi=typeof(Model).GetProperty("Value");
__P((pi.CanRead).ToString()); __P((pi.CanWrite).ToString());
__Check("True\nTrue");
