// vybe-test: csharp/csharp_constructor_chains/record_primary_constructor_members_are_available_immediately
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User(string Name, int Age); var user = new User("Ada", 20); __Check((user.Name).ToString(), "Ada"); __Check((user.Age).ToString(), "20");
