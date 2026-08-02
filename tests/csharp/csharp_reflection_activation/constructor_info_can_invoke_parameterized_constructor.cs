// vybe-test: csharp/csharp_reflection_activation/constructor_info_can_invoke_parameterized_constructor
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public string Name; public Box(string name) { Name = name; } } var ctor = typeof(Box).GetConstructor(new[] { typeof(string) }); var box = (Box)ctor.Invoke(new object[] { "crate" }); __Check((box.Name).ToString(), "crate");
