// vybe-test: csharp/csharp_record_struct_deep/record_struct_positional_init_with
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct User(string Name){public int Age{get;init;}} var u=new User("Ada"){Age=30}; var v=u with{Age=31}; __Check((u.Age).ToString(), "30"); __Check((v.Age).ToString(), "31");
