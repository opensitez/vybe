// vybe-test: csharp/csharp_events/null_event_invocation_via_conditional_access_is_safe
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Source { public event System.Action Fired; public void Fire() => Fired?.Invoke(); }
var s = new Source();
s.Fire();
__Check(("ok").ToString(), "ok");
