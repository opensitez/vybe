// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_pointer_equality_same_element
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

using static __Harness;

byte[] arr = new byte[] { 10, 20, 30, 40 };
Span<byte> ptr = arr.AsSpan();
__P(ptr[0].ToString());
__P(ptr[1].ToString());
__Check("10\n20");

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
