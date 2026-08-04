// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

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

// serialization_json_surface
string feature = "serialization_json_surface"; __P((feature.Substring(0, 1) == feature[0].ToString()).ToString());
__Check("True");
