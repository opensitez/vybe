// vybe-test: csharp/csharp_custom_event_accessors/custom_event_eventhandler_generic
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Msg: System.EventArgs{public string Text;} class Ch{System.EventHandler<Msg> _h; public event System.EventHandler<Msg> Sent{add{_h+=value;} remove{_h-=value;}} public void Emit(string t){_h?.Invoke(this,new Msg{Text=t});}} string out_=""; var c=new Ch(); c.Sent+=(o,e)=>out_=e.Text; c.Emit("hi"); __Check((out_).ToString(), "hi");
