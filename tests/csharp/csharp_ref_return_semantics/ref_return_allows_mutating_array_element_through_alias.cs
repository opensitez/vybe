// vybe-test: csharp/csharp_ref_return_semantics/ref_return_allows_mutating_array_element_through_alias
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

using static __Harness;

int[] data = { 1, 2, 3 }
;
ref int Slot(int index) => ref data[index];
ref int cell = ref Slot(1);
cell = 9;
__P((data[1]).ToString());
__Check("9");

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
