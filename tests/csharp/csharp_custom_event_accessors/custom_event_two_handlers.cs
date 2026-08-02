// vybe-test: csharp/csharp_custom_event_accessors/custom_event_two_handlers
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}var log=""; var b=new Btn(); b.Click+=()=>log+="a"; b.Click+=()=>log+="b"; b.Raise(); __Check((log).ToString(), "ab");
