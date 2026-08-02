// vybe-test: csharp/csharp_reflection_activation/method_info_invokes_instance_method_and_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public string Read() { return "value"; } } var method = typeof(Box).GetMethod("Read"); __Check((method.Invoke(new Box(), null)).ToString(), "value");
