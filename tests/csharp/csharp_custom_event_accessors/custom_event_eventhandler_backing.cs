// vybe-test: csharp/csharp_custom_event_accessors/custom_event_eventhandler_backing
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

class Src{System.EventHandler _h; public event System.EventHandler Changed{add{_h+=value;} remove{_h-=value;}} public void Raise(){_h?.Invoke(this,System.EventArgs.Empty);}} int n=0; var s=new Src(); s.Changed+=(o,e)=>n++; s.Raise(); __P((n).ToString());
__Check("1");
