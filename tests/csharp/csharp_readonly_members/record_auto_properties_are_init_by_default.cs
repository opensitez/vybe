// vybe-test: csharp/csharp_readonly_members/record_auto_properties_are_init_by_default
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

record User(string Name,int Age);
var u=new User("Ada",20);
__P((u.Name).ToString()); __P((u.Age).ToString());
__Check("Ada\n20");
