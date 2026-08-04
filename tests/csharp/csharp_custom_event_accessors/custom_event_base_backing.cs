// vybe-test: csharp/csharp_custom_event_accessors/custom_event_base_backing
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

class Base{System.Action _e; public event System.Action Ping{add{_e+=value;} remove{_e-=value;}} protected void OnPing(){_e?.Invoke();}} class Child:Base{public void Fire(){OnPing();}} int n=0; var c=new Child(); c.Ping+=()=>n++; c.Fire(); __P((n).ToString());
__Check("1");
