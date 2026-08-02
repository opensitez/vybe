// vybe-test: csharp/csharp_custom_event_accessors/custom_event_sender_passed_to_handler
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Src{System.EventHandler _h; public event System.EventHandler Tick{add{_h+=value;} remove{_h-=value;}} public void Pulse(){_h?.Invoke(this,System.EventArgs.Empty);}} object who=null; var s=new Src(); s.Tick+=(sender,e)=>who=sender; s.Pulse(); __Check((who==s).ToString(), "True");
