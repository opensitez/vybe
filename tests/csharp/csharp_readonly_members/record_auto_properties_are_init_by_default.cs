// vybe-test: csharp/csharp_readonly_members/record_auto_properties_are_init_by_default
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User(string Name,int Age);
var u=new User("Ada",20);
__Check((u.Name).ToString(), "Ada"); __Check((u.Age).ToString(), "20");
