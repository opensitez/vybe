// vybe-test: csharp/csharp_event_advanced/event_handler_generic_passes_custom_event_args
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

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

class ChangedArgs:System.EventArgs{public string OldValue;public string NewValue;}
class Setting{
    private string _v="";
    public event System.EventHandler<ChangedArgs> Changed;
    public string Value{
        get=>_v;
        set{var old=_v;_v=value;Changed?.Invoke(this,new ChangedArgs{OldValue=old,NewValue=_v});}
    }
}
var s=new Setting();
string log="";
s.Changed+=(o,e)=>log=$"{e.OldValue}->{e.NewValue}";
s.Value="hello";
__P((log).ToString());
__Check("->hello");
