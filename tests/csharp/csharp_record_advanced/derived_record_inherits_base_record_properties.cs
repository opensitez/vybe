// vybe-test: csharp/csharp_record_advanced/derived_record_inherits_base_record_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

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

record Animal(string Name);
record Dog(string Name,string Breed):Animal(Name);
var d=new Dog("Rex","Lab");
__P((d.Name).ToString()); __P((d.Breed).ToString());
__Check("Rex\nLab");
