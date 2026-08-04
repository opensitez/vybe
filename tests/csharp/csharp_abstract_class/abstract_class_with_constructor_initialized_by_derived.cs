// vybe-test: csharp/csharp_abstract_class/abstract_class_with_constructor_initialized_by_derived
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

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

abstract class Named{public string Name;public Named(string n){Name=n;}}
class Tag:Named{public Tag(string n):base(n){}}
__P((new Tag("admin").Name).ToString());
__Check("admin");
