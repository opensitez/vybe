// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_delegate_can_be_invoked_through_qualified_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo { public delegate string Reader(); } Demo.Reader reader = () => "text"; __Check((reader()).ToString(), "text");
