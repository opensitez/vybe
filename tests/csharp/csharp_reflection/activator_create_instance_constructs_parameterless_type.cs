// vybe-test: csharp/csharp_reflection/activator_create_instance_constructs_parameterless_type
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget { public int Value = 42; }
var w = (Widget)System.Activator.CreateInstance(typeof(Widget));
__Check((w.Value).ToString(), "42");
