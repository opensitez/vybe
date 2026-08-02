// vybe-test: csharp/csharp_custom_event_accessors/custom_event_has_after_remove
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public bool Has=>_c!=null; public void Raise(){_c?.Invoke();}} System.Action h=()=>{}; var b=new Btn(); b.Click+=h; b.Click-=h; __Check((b.Has).ToString(), "False");
