// vybe-test: csharp/csharp_generics_advanced/generic_where_new_constraint_creates_instance_inside_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget { public int Val = 5; }
T Make<T>() where T : new() => new T();
__Check((Make<Widget>().Val).ToString(), "5");
