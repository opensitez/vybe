// vybe-test: csharp/csharp_yield_advanced/yield_return_with_complex_state_machine
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

using static __Harness;

System.Collections.Generic.IEnumerable<string> Words(string s){
    var parts=s.Split(' ');
    foreach(var p in parts) if(p.Length>0) yield return p;
}
__P((string.Join("|",Words("hello  world  foo"))).ToString());
__Check("hello|world|foo");

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
