// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
    int _v = 5;
    public int Value {
        get {
            Report();
            return _v;
        }
    }
    void Report([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Value");
}
__Check((new Box().Value).ToString(), "5");
