// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_on_static_property_setter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config {
    static int _port;
    public static int Port {
        set {
            Log();
            _port = value;
        }
        get => _port;
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Port");
}
Config.Port = 443; __Check((Config.Port).ToString(), "443");
