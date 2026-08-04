// vybe-test: csharp/csharp_record_types/record_hash_code_equal_for_equal_instances
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

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

record Tag(string Name);
var a = new Tag("x"); var b = new Tag("x");
__P((a.GetHashCode() == b.GetHashCode()).ToString());
__Check("True");
