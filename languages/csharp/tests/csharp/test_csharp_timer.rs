//! `System.Threading.Timer` and `System.Timers.Timer` basics.
use super::helpers::run_csharp;

#[test]
fn threading_timer_fires_callback_after_delay() {
    assert_eq!(
        run_csharp(
            r#"bool fired=false;
using var t=new System.Threading.Timer(_=>{fired=true;},null,10,System.Threading.Timeout.Infinite);
System.Threading.Thread.Sleep(100);
Console.WriteLine(fired);"#
        ),
        &["True"]
    );
}

#[test]
fn timers_timer_elapsed_event_fires() {
    assert_eq!(
        run_csharp(
            r#"bool fired=false;
var t=new System.Timers.Timer(10){AutoReset=false};
t.Elapsed+=(_,__)=>fired=true;
t.Start();
System.Threading.Thread.Sleep(100);
Console.WriteLine(fired);"#
        ),
        &["True"]
    );
}

#[test]
fn timer_change_reschedules_callback() {
    assert_eq!(
        run_csharp(
            r#"int count=0;
using var t=new System.Threading.Timer(_=>System.Threading.Interlocked.Increment(ref count),null,10,10);
System.Threading.Thread.Sleep(100);
Console.WriteLine(count>0);"#
        ),
        &["True"]
    );
}
