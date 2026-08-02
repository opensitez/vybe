// vybe-test: csharp/csharp_event_advanced/event_handler_generic_passes_custom_event_args
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((log).ToString(), "->hello");
