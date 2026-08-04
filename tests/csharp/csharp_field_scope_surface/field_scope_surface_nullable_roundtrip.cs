// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

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

// field_scope_surface
int? maybe = 63; __P((maybe.HasValue && maybe.Value == 63).ToString());
__Check("True");
