// vybe-test: csharp/csharp_struct_advanced/readonly_struct_method_cannot_modify_state
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Counter{
    public readonly int Value;
    public Counter(int v){Value=v;}
    public Counter Increment()=>new Counter(Value+1);
}
var c=new Counter(5).Increment();
__Check((c.Value).ToString(), "6");
