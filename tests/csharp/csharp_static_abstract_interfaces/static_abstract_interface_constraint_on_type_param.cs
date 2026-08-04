// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_interface_constraint_on_type_param
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

interface IHasLabel<T> where T:IHasLabel<T>{static abstract string Label();}
struct Tag:IHasLabel<Tag>{public static string Label()=>"tag";}
string Read<T>() where T:IHasLabel<T>=>T.Label(); __P((Read<Tag>()).ToString());
__Check("tag");
