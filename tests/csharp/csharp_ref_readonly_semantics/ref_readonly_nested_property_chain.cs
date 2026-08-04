// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_nested_property_chain
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

class Node{public int Value;} class Holder{public Node Inner=new Node();} var h=new Holder(); h.Inner.Value=33; ref readonly int Read(ref Holder host)=>ref host.Inner.Value; __P((Read(ref h)).ToString());
__Check("33");
