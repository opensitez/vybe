// vybe-test: csharp/csharp_custom_event_accessors/custom_event_public_subscribe_private_raise
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

class Hub{System.Action _h; public event System.Action Signal{add{_h+=value;} remove{_h-=value;}} public void Pulse(){_h?.Invoke();}} int v=0; var h=new Hub(); h.Signal+=()=>v=9; h.Pulse(); __P((v).ToString());
__Check("9");
