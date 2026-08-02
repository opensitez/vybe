// vybe-test: csharp/csharp_reflection/get_methods_count_includes_public_instance_methods
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Calc { public int Add(int a, int b) => a+b; public int Sub(int a, int b) => a-b; }
__Check((typeof(Calc).GetMethods(
    System.Reflection.BindingFlags.Public|System.Reflection.BindingFlags.Instance|
    System.Reflection.BindingFlags.DeclaredOnly).Length).ToString(), "2");
