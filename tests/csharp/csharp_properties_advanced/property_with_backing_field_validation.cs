// vybe-test: csharp/csharp_properties_advanced/property_with_backing_field_validation
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

class Age{
    int _value;
    public int Value{
        get=>_value;
        set{if(value<0)throw new System.ArgumentException();_value=value;}
    }
}
var a=new Age{Value=25};
__P((a.Value).ToString());
__Check("25");
