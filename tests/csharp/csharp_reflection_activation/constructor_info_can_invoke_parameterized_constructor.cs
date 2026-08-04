// vybe-test: csharp/csharp_reflection_activation/constructor_info_can_invoke_parameterized_constructor
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

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

using System; class Box { public string Name; public Box(string name) { Name = name; } } var ctor = typeof(Box).GetConstructor(new[] { typeof(string) }); var box = (Box)ctor.Invoke(new object[] { "crate" }); __P((box.Name).ToString());
__Check("crate");
