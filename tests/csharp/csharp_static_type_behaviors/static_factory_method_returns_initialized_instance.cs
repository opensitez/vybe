// vybe-test: csharp/csharp_static_type_behaviors/static_factory_method_returns_initialized_instance
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class User {
    public string Name { get; set; }
    public static User CreateAdmin() { return new User { Name = "root" }; }
}
var user = User.CreateAdmin();
__P((user.Name).ToString());
__Check("root");
