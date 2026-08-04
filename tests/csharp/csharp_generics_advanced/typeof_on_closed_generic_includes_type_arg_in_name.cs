// vybe-test: csharp/csharp_generics_advanced/typeof_on_closed_generic_includes_type_arg_in_name
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

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

__P((typeof(System.Collections.Generic.List<int>).IsGenericType).ToString());
__Check("True");
