// vybe-test: csharp/csharp_static_type_behaviors/generic_static_field_is_scoped_per_closed_type
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

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

class Cache<T> {
    public static int Hits;
}
Cache<int>.Hits++;
Cache<int>.Hits++;
Cache<string>.Hits++;
__P((Cache<int>.Hits).ToString());
__P((Cache<string>.Hits).ToString());
__Check("2\n1");
