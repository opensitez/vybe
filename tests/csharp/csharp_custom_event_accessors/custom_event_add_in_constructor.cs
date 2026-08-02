// vybe-test: csharp/csharp_custom_event_accessors/custom_event_add_in_constructor
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public Btn(){_c+=()=>_boot=1;} int _boot; public void Raise(){_c?.Invoke();} public int Boot=>_boot;} var b=new Btn(); __Check((b.Boot).ToString(), "1");
