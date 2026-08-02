// vybe-test: csharp/csharp_static_type_behaviors/static_field_is_shared_across_instances
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Session {
    public static int Count = 0;
    public Session() { Count++; }
}
new Session();
new Session();
__Check((Session.Count).ToString(), "2");
