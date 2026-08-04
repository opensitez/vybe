// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_default_string_label
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

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

interface ILabel<T> where T:ILabel<T>{static virtual string Tag=>"d"; static abstract T Make();}
struct Tag:ILabel<Tag>{public static Tag Make()=>new Tag(); public static string Tag=>"x";}
__P((Tag.Tag).ToString());
__Check("x");
