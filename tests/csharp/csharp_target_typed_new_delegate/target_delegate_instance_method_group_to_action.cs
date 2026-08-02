// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_instance_method_group_to_action
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Logger { public string last = ""; public void Save(string msg) => last = msg; }
var log = new Logger();
System.Action<string> write = log.Save;
write("note");
__Check((log.last).ToString(), "note");
