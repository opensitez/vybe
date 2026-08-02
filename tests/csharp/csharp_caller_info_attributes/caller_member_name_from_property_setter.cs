// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_property_setter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
    int _v;
    public int Value {
        set {
            Report();
            _v = value;
        }
        get => _v;
    }
    void Report([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Value");
}
var b = new Box(); b.Value = 9; __Check((b.Value).ToString(), "9");
