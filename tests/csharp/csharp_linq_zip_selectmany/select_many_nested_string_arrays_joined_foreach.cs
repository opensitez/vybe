// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_nested_string_arrays_joined_foreach
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var flat=new[]{new[]{"a","b"},new[]{"c"}}
.SelectMany(x=>x);
foreach(var s in flat) __P((s).ToString());
__Check("a\nb\nc");

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
