// vybe-test: csharp/csharp_custom_event_accessors/custom_event_wrong_instance_unsub
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} int n=0; System.Action h=()=>n++; var a=new Btn(); var b=new Btn(); a.Click+=h; b.Click-=h; a.Raise(); __Check((n).ToString(), "1");
