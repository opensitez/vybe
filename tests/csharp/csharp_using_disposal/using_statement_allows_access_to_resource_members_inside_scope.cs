// vybe-test: csharp/csharp_using_disposal/using_statement_allows_access_to_resource_members_inside_scope
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public string Name => "file"; public void Dispose() { __Check(("disposed").ToString(), "file"); } } using (var resource = new Resource()) { __Check((resource.Name).ToString(), "disposed"); }
