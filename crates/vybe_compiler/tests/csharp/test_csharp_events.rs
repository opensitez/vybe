//! Event declaration, subscription, unsubscription, and `EventArgs`.
use super::helpers::run_csharp;

#[test]
fn event_fires_single_subscriber() {
    assert_eq!(
        run_csharp(
            r#"class Button {
    public event System.EventHandler Clicked;
    public void Click() => Clicked?.Invoke(this, System.EventArgs.Empty);
}
int count = 0;
var btn = new Button();
btn.Clicked += (s, e) => count++;
btn.Click();
Console.WriteLine(count);"#
        ),
        &["1"]
    );
}

#[test]
fn unsubscribed_handler_not_called_after_removal() {
    assert_eq!(
        run_csharp(
            r#"class Button {
    public event System.EventHandler Clicked;
    public void Click() => Clicked?.Invoke(this, System.EventArgs.Empty);
}
int count = 0;
System.EventHandler h = (s, e) => count++;
var btn = new Button();
btn.Clicked += h;
btn.Clicked -= h;
btn.Click();
Console.WriteLine(count);"#
        ),
        &["0"]
    );
}

#[test]
fn multiple_subscribers_all_called_in_order() {
    assert_eq!(
        run_csharp(
            r#"class Emitter {
    public event System.Action<string> Signal;
    public void Emit(string v) => Signal?.Invoke(v);
}
string log = "";
var e = new Emitter();
e.Signal += v => log += "A";
e.Signal += v => log += "B";
e.Emit("x");
Console.WriteLine(log);"#
        ),
        &["AB"]
    );
}

#[test]
fn custom_event_args_carries_data_to_handler() {
    assert_eq!(
        run_csharp(
            r#"class DataArgs : System.EventArgs { public int Value; }
class Source {
    public event System.EventHandler<DataArgs> Changed;
    public void Change(int v) => Changed?.Invoke(this, new DataArgs{Value=v});
}
int received = 0;
var src = new Source();
src.Changed += (s, e) => received = e.Value;
src.Change(77);
Console.WriteLine(received);"#
        ),
        &["77"]
    );
}

#[test]
fn null_event_invocation_via_conditional_access_is_safe() {
    assert_eq!(
        run_csharp(
            r#"class Source { public event System.Action Fired; public void Fire() => Fired?.Invoke(); }
var s = new Source();
s.Fire();
Console.WriteLine("ok");"#
        ),
        &["ok"]
    );
}
