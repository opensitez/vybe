// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

using static __Harness;

// dictionary_insert_lookup
int? maybe = 34;
__P((maybe.HasValue && maybe.Value == 34).ToString());
__Check("True");

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
