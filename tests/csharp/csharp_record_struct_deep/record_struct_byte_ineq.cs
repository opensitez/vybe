// vybe-test: csharp/csharp_record_struct_deep/record_struct_byte_ineq
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

__P((new ByteVal(1)==new ByteVal(2)).ToString());
__Check("False");

record struct ByteVal(byte B);

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
