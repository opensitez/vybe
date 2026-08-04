// vybe-test: csharp/csharp_custom_event_accessors/custom_event_count_tracked
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

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

class Btn{System.Action _e; int _count; public event System.Action Tick{add{_e+=value;_count++;} remove{_e-=value;_count--;}} public int Count=>_count; public void Fire(){_e?.Invoke();}} var b=new Btn(); System.Action h=()=>{}; b.Tick+=h; b.Tick+=()=>{}; b.Tick-=h; __P((b.Count).ToString());
__Check("1");
