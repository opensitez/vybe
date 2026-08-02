// vybe-test: csharp/csharp_custom_event_accessors/custom_event_base_backing
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base{System.Action _e; public event System.Action Ping{add{_e+=value;} remove{_e-=value;}} protected void OnPing(){_e?.Invoke();}} class Child:Base{public void Fire(){OnPing();}} int n=0; var c=new Child(); c.Ping+=()=>n++; c.Fire(); __Check((n).ToString(), "1");
