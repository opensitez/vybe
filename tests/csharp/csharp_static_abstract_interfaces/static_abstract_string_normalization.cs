// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_string_normalization
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

interface INorm<T> where T:INorm<T>{static abstract T Normalize(string s);}
struct Text:INorm<Text>{public string Value; public static Text Normalize(string s)=>new Text{Value=s.Trim().ToLower()};}
__P((Text.Normalize(" Ab ").Value).ToString());
__Check("ab");
