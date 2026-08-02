// vybe-test: csharp/csharp_access_modifiers/public_method_callable_from_any_scope
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Service{public string Name()=>"svc";}
__Check((new Service().Name()).ToString(), "svc");
