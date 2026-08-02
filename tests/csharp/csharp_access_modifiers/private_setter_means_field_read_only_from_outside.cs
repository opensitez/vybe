// vybe-test: csharp/csharp_access_modifiers/private_setter_means_field_read_only_from_outside
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter{
    public int Count{get;private set;}
    public void Tick(){Count++;}
}
var c=new Counter(); c.Tick(); c.Tick();
__Check((c.Count).ToString(), "2");
