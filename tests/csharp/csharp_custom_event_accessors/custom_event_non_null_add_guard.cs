// vybe-test: csharp/csharp_custom_event_accessors/custom_event_non_null_add_guard
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{if(value!=null)_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Raise(); __Check((n).ToString(), "1");
