// vybe-test: csharp/csharp_oop_advanced2/abstract_class_partial_implementation_forces_concrete_override
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Step{
    public abstract string Name();
    public string Run()=>$"run:{Name()}";
}
class Alpha:Step{public override string Name()=>"alpha";}
__Check((new Alpha().Run()).ToString(), "run:alpha");
