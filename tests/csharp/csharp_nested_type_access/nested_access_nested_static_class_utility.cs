// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_static_class_utility
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

class Text{public static class Util{public static string Join(string a,string b)=>a+b;} public static string Merge()=>Util.Join("a","b");} __P((Text.Merge()).ToString());
__Check("ab");
