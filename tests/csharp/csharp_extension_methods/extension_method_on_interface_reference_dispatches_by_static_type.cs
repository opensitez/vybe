// vybe-test: csharp/csharp_extension_methods/extension_method_on_interface_reference_dispatches_by_static_type
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; interface IName { string Name { get; } } class User : IName { public string Name => "Ada"; } namespace Demo { public static class NameExt { public static string UpperName(this IName value) { return value.Name.ToUpper(); } } } IName user = new User(); __Check((user.UpperName()).ToString(), "ADA");
