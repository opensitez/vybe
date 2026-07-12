//! Events: custom accessor, thread-safe, EventHandler<T>, unsubscribe.
use super::helpers::run_csharp;

#[test]
fn event_handler_generic_passes_custom_event_args() {
    assert_eq!(
        run_csharp(
            r#"class ChangedArgs:System.EventArgs{public string OldValue;public string NewValue;}
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
Console.WriteLine(log);"#
        ),
        &["->hello"]
    );
}

#[test]
fn unsubscribing_from_event_stops_handler_firing() {
    assert_eq!(
        run_csharp(
            r#"class Btn{public event System.Action Click;}
int count=0;
System.Action h=()=>count++;
var b=new Btn();
b.Click+=h;
b.Click?.Invoke();
b.Click-=h;
b.Click?.Invoke();
Console.WriteLine(count);"#
        ),
        &["1"]
    );
}

#[test]
fn null_conditional_event_invoke_safe_when_no_subscribers() {
    assert_eq!(
        run_csharp(
            r#"class Btn{public event System.Action Click;}
var b=new Btn();
b.Click?.Invoke();
Console.WriteLine("ok");"#
        ),
        &["ok"]
    );
}

#[test]
fn multicast_event_all_subscribers_called_in_order() {
    assert_eq!(
        run_csharp(
            r#"class Btn{public event System.Action Click;}
var b=new Btn();
var log=new System.Collections.Generic.List<string>();
b.Click+=()=>log.Add("a");
b.Click+=()=>log.Add("b");
b.Click?.Invoke();
Console.WriteLine(string.Join(",",log));"#
        ),
        &["a,b"]
    );
}
