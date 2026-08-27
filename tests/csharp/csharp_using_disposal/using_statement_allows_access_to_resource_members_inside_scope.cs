// vybe-test: csharp/csharp_using_disposal/using_statement_allows_access_to_resource_members_inside_scope
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;
using System;

using (var resource = new Resource()) { __P((resource.Name).ToString()); }
__Check("file\ndisposed");

class Resource : IDisposable { public string Name => "file"; public void Dispose() { __P(("disposed").ToString()); } }

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
