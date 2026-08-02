// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_action_string_from_method_group
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string last = "";
void Capture(string s) { last = s; }
System.Action<string> store = Capture;
store("saved");
__Check((last).ToString(), "saved");
