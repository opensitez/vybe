// vybe-test: csharp/csharp_reflection_activation/field_info_sets_public_field_value_on_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public int Count; } var box = new Box(); var field = typeof(Box).GetField("Count"); field.SetValue(box, 9); __Check((box.Count).ToString(), "9");
