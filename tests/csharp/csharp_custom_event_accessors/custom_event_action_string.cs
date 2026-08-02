// vybe-test: csharp/csharp_custom_event_accessors/custom_event_action_string
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Line{System.Action<string> _h; public event System.Action<string> Write{add{_h+=value;} remove{_h-=value;}} public void Emit(string s){_h?.Invoke(s);}} string log=""; var l=new Line(); l.Write+=s=>log+=s; l.Emit("x"); __Check((log).ToString(), "x");
