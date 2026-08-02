// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_read_in_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter(int start) {
    int current = start;
    public int Next() => ++current;
    public int Value => current;
}
var c = new Counter(10);
c.Next(); c.Next();
__Check((c.Value).ToString(), "12");
