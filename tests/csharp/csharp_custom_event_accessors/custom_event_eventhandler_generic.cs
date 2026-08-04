// vybe-test: csharp/csharp_custom_event_accessors/custom_event_eventhandler_generic
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

class Msg: System.EventArgs{public string Text;} class Ch{System.EventHandler<Msg> _h; public event System.EventHandler<Msg> Sent{add{_h+=value;} remove{_h-=value;}} public void Emit(string t){_h?.Invoke(this,new Msg{Text=t});}} string out_=""; var c=new Ch(); c.Sent+=(o,e)=>out_=e.Text; c.Emit("hi"); __P((out_).ToString());
__Check("hi");
