// vybe-test: csharp/csharp_constructor_patterns/primary_constructor_on_record_sets_all_fields
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

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

record Person(string Name,int Age);
var p=new Person("Grace",40);
__P((p.Name).ToString()); __P((p.Age).ToString());
__Check("Grace\n40");
