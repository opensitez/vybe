// vybe-test: csharp/csharp_properties_advanced/static_property_shared_across_instances
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

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

class AppConfig{public static string Version{get;set;}="1.0";}
AppConfig.Version="2.0";
__P((new System.Object().GetType()!=null).ToString());
__P((AppConfig.Version).ToString());
__Check("True\n2.0");
