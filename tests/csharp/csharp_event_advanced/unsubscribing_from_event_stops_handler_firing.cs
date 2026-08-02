// vybe-test: csharp/csharp_event_advanced/unsubscribing_from_event_stops_handler_firing
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Btn{public event System.Action Click;}
int count=0;
System.Action h=()=>count++;
var b=new Btn();
b.Click+=h;
b.Click?.Invoke();
b.Click-=h;
b.Click?.Invoke();
__Check((count).ToString(), "1");
