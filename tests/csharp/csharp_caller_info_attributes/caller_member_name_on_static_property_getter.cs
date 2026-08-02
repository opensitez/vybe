// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_on_static_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config {
    static int _port = 80;
    public static int Port {
        get {
            Log();
            return _port;
        }
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Port");
}
__Check((Config.Port).ToString(), "80");
