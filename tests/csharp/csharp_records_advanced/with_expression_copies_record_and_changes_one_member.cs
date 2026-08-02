// vybe-test: csharp/csharp_records_advanced/with_expression_copies_record_and_changes_one_member
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User(string Name, int Age); var user = new User("Ada", 20); var updated = user with { Age = 21 }; __Check((user.Age).ToString(), "20"); __Check((updated.Age).ToString(), "21");
