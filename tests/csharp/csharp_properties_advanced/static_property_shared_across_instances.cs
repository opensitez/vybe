// vybe-test: csharp/csharp_properties_advanced/static_property_shared_across_instances
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class AppConfig{public static string Version{get;set;}="1.0";}
AppConfig.Version="2.0";
__Check((new System.Object().GetType()!=null).ToString(), "True");
__Check((AppConfig.Version).ToString(), "2.0");
