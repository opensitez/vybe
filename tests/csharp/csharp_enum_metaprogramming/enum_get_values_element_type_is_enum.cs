// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_element_type_is_enum
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

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

enum Kind{A,B} var values=System.Enum.GetValues(typeof(Kind)); __P((values.GetType().GetElementType().Name).ToString());
__Check("Kind");
