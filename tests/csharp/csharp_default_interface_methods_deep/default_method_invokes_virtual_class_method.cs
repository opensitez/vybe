// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_invokes_virtual_class_method
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRun{string Go(){return Run();} string Run();} class Job:IRun{public string Run()=>"done";} __Check((new Job().Go()).ToString(), "done");
