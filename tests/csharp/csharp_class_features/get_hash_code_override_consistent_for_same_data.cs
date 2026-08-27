// vybe-test: csharp/csharp_class_features/get_hash_code_override_consistent_for_same_data
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

using static __Harness;

__P((new Key(7).GetHashCode()==new Key(7).GetHashCode()).ToString());
__Check("True");

class Key{int V;public Key(int v){V=v;}public override int GetHashCode()=>V.GetHashCode();}

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
