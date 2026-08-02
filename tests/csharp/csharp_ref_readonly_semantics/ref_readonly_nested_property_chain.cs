// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_nested_property_chain
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Node{public int Value;} class Holder{public Node Inner=new Node();} var h=new Holder(); h.Inner.Value=33; ref readonly int Read(ref Holder host)=>ref host.Inner.Value; __Check((Read(ref h)).ToString(), "33");
