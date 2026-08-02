// vybe-test: csharp/csharp_custom_event_accessors/custom_event_eventhandler_backing
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Src{System.EventHandler _h; public event System.EventHandler Changed{add{_h+=value;} remove{_h-=value;}} public void Raise(){_h?.Invoke(this,System.EventArgs.Empty);}} int n=0; var s=new Src(); s.Changed+=(o,e)=>n++; s.Raise(); __Check((n).ToString(), "1");
