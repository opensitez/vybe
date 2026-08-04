// vybe-test: csharp/csharp_reflection_activation/field_info_sets_public_field_value_on_instance
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

using System; class Box { public int Count; } var box = new Box(); var field = typeof(Box).GetField("Count"); field.SetValue(box, 9); __P((box.Count).ToString());
__Check("9");
