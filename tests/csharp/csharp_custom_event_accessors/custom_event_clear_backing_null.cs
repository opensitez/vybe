// vybe-test: csharp/csharp_custom_event_accessors/custom_event_clear_backing_null
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Clear(){_c=null;} public bool Empty=>_c==null;} var b=new Btn(); b.Click+=()=>{}; b.Clear(); __Check((b.Empty).ToString(), "True");
