// vybe-test: csharp/csharp_extension_methods/extension_method_can_use_namespace_import_from_nested_namespace
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using Demo.Tools; namespace Demo.Tools { public static class TextExt { public static string Bang(this string value) { return value + "!"; } } } __P(("go".Bang()).ToString());
__Check("go!");
