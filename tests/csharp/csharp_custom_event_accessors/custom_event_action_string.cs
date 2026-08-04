// vybe-test: csharp/csharp_custom_event_accessors/custom_event_action_string
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

class Line{System.Action<string> _h; public event System.Action<string> Write{add{_h+=value;} remove{_h-=value;}} public void Emit(string s){_h?.Invoke(s);}} string log=""; var l=new Line(); l.Write+=s=>log+=s; l.Emit("x"); __P((log).ToString());
__Check("x");
