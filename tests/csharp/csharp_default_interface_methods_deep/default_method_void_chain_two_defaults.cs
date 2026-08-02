// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_void_chain_two_defaults
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IA{void A(){__Check(("a").ToString(), "a");}} interface IB{void B(){__Check(("b").ToString(), "b");}} class Both:IA,IB{} var b=new Both(); b.A(); b.B();
