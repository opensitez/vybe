// vybe-test: csharp/csharp_null_propagation/null_coalescing_can_select_new_object_instance
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public string Name; } Box box = null; box ??= new Box { Name = "created" }; __Check((box.Name).ToString(), "created");
