// vybe-test: csharp/csharp_anonymous_types/anonymous_type_property_names_inferred_from_variable
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

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

int id=7; string name="Bob";
var obj=new{id,name};
__P((obj.id).ToString()); __P((obj.name).ToString());
__Check("7\nBob");
