// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_on_static_property_setter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

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

class Config {
    static int _port;
    public static int Port {
        set {
            Log();
            _port = value;
        }
        get => _port;
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __P((member).ToString());
}
Config.Port = 443; __P((Config.Port).ToString());
__Check("Port\n443");
