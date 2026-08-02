// vybe-test: csharp/csharp_default_interface_methods_deep/default_void_method_with_side_effect
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ILog{void Ping(){__Check(("ping").ToString(), "ping");}} class Silent:ILog{} new Silent().Ping();
