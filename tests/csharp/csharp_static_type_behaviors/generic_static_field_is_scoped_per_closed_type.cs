// vybe-test: csharp/csharp_static_type_behaviors/generic_static_field_is_scoped_per_closed_type
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

using static __Harness;

Cache<int>.Hits++;
Cache<int>.Hits++;
Cache<string>.Hits++;
__P((Cache<int>.Hits).ToString());
__P((Cache<string>.Hits).ToString());
__Check("2\n1");

class Cache<T> {
    public static int Hits;
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
