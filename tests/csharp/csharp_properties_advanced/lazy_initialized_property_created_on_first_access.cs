// vybe-test: csharp/csharp_properties_advanced/lazy_initialized_property_created_on_first_access
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

class Config{
    System.Lazy<string> _tag=new System.Lazy<string>(()=>"computed");
    public string Tag=>_tag.Value;
}
var c=new Config();
__P((c.Tag).ToString());
__Check("computed");
