// vybe-test: csharp/csharp_properties_advanced/lazy_initialized_property_created_on_first_access
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config{
    System.Lazy<string> _tag=new System.Lazy<string>(()=>"computed");
    public string Tag=>_tag.Value;
}
var c=new Config();
__Check((c.Tag).ToString(), "computed");
