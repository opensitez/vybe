// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_instance_method_group_to_action
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

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

class Logger { public string last = ""; public void Save(string msg) => last = msg; }
var log = new Logger();
System.Action<string> write = log.Save;
write("note");
__P((log.last).ToString());
__Check("note");
