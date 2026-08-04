// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_arithmetic_zeroing
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
int seed = 91; __P((seed - seed == 0).ToString());
__Check("True");
