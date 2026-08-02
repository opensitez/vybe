// vybe-test: csharp/csharp_virtual_dispatch_semantics/base_keyword_invokes_parent_implementation_from_override
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public virtual int Next() { return 1; }
}
class StepCounter : Counter {
    public override int Next() { return base.Next() + 2; }
}
__Check((new StepCounter().Next()).ToString(), "3");
