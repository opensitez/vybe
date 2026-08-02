// vybe-test: csharp/csharp_properties_advanced/property_with_backing_field_validation
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Age{
    int _value;
    public int Value{
        get=>_value;
        set{if(value<0)throw new System.ArgumentException();_value=value;}
    }
}
var a=new Age{Value=25};
__Check((a.Value).ToString(), "25");
