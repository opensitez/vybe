// vybe-test: csharp/csharp_custom_event_accessors/custom_event_public_subscribe_private_raise
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Hub{System.Action _h; public event System.Action Signal{add{_h+=value;} remove{_h-=value;}} public void Pulse(){_h?.Invoke();}} int v=0; var h=new Hub(); h.Signal+=()=>v=9; h.Pulse(); __Check((v).ToString(), "9");
