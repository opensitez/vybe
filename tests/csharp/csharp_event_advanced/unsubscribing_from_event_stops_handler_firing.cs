// vybe-test: csharp/csharp_event_advanced/unsubscribing_from_event_stops_handler_firing
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((count).ToString());
__Check("1");
