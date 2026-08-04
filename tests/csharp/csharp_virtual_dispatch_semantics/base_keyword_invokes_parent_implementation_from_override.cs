// vybe-test: csharp/csharp_virtual_dispatch_semantics/base_keyword_invokes_parent_implementation_from_override
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public virtual int Next() { return 1; }
}
class StepCounter : Counter {
    public override int Next() { return base.Next() + 2; }
}
__P((new StepCounter().Next()).ToString());
__Check("3");
