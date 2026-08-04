// vybe-test: csharp/csharp_readonly_members/init_property_settable_only_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

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

class Config{public int Port{get;init;}=80;}
var c=new Config{Port=443};
__P((c.Port).ToString());
__Check("443");
