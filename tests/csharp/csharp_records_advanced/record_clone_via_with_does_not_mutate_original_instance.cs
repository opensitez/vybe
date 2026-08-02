// vybe-test: csharp/csharp_records_advanced/record_clone_via_with_does_not_mutate_original_instance
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User(string Name, int Age); var before = new User("Ada", 30); var after = before with { Name = "Grace" }; __Check((before.Name).ToString(), "Ada"); __Check((after.Name).ToString(), "Grace");
