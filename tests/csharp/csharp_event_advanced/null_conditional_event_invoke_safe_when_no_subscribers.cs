// vybe-test: csharp/csharp_event_advanced/null_conditional_event_invoke_safe_when_no_subscribers
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{public event System.Action Click;}
var b=new Btn();
b.Click?.Invoke();
__Check(("ok").ToString(), "ok");
