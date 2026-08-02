// vybe-test: csharp/csharp_custom_event_accessors/custom_event_lock_in_accessor
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; readonly object _gate=new object(); public event System.Action Click{add{lock(_gate){_c+=value;}} remove{lock(_gate){_c-=value;}}} public void Raise(){_c?.Invoke();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Raise(); __Check((n).ToString(), "1");
