// vybe-test: csharp/csharp_nested_type_access/nested_access_private_nested_not_exposed_but_outer_delegates
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Service{class Engine{public string Run()=>"ok";} public string Execute()=>new Engine().Run();} __Check((new Service().Execute()).ToString(), "ok");
