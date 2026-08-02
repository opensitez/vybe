// vybe-test: csharp/csharp_reflection_activation/activator_create_instance_builds_default_constructed_object
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public int Value = 4; } var box = (Box)Activator.CreateInstance(typeof(Box)); __Check((box.Value).ToString(), "4");
