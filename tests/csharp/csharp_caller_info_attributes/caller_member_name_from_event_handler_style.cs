// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_event_handler_style
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn {
    public void Click() { OnClick(); }
    void OnClick([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Click");
}
new Btn().Click();
