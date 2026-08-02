// vybe-test: csharp/csharp_static_type_behaviors/static_factory_method_returns_initialized_instance
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class User {
    public string Name { get; set; }
    public static User CreateAdmin() { return new User { Name = "root" }; }
}
var user = User.CreateAdmin();
__Check((user.Name).ToString(), "root");
