// vybe-test: csharp/csharp_reflection_activation/field_info_reads_public_field_value_from_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public int Count = 12; } var field = typeof(Box).GetField("Count"); __Check((field.GetValue(new Box())).ToString(), "12");
