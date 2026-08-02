// vybe-test: csharp/csharp_reflection_emit/property_info_get_set_accessor_names
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Model{public int Value{get;set;}}
var pi=typeof(Model).GetProperty("Value");
__Check((pi.CanRead).ToString(), "True"); __Check((pi.CanWrite).ToString(), "True");
