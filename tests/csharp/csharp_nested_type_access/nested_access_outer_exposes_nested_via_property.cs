// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_exposes_nested_via_property
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Shell{public class Core{public int Id=2;} Core _c=new Core(); public Core Inner=>_c;} __P((new Shell().Inner.Id).ToString());
__Check("2");
