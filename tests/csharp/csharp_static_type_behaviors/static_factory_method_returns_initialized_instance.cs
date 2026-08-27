// vybe-test: csharp/csharp_static_type_behaviors/static_factory_method_returns_initialized_instance
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

using static __Harness;

var user = User.CreateAdmin();
__P((user.Name).ToString());
__Check("root");

class User {
    public string Name { get; set; }
    public static User CreateAdmin() { return new User { Name = "root" }; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
