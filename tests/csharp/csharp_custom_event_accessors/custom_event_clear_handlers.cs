// vybe-test: csharp/csharp_custom_event_accessors/custom_event_clear_handlers
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

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

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();} public int Count=>_c==null?0:_c.GetInvocationList().Length;} System.Action a=()=>{}; System.Action b=()=>{}; var btn=new Btn(); btn.Click+=a; btn.Click+=b; btn.Click-=a; btn.Click-=b; __P((btn.Count).ToString());
__Check("0");
