// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_passed_to_static_helper
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Worker(int id) {
    static string Format(int value) => "id=" + value;
    public string Show() => Format(id);
}
__Check((new Worker(3).Show()).ToString(), "id=3");
