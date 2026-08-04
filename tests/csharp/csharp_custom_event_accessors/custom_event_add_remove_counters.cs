// vybe-test: csharp/csharp_custom_event_accessors/custom_event_add_remove_counters
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

class Btn{System.Action _c; public int Adds=0; public int Removes=0; public event System.Action Click{add{Adds++; _c+=value;} remove{Removes++; _c-=value;}} public void Raise(){_c?.Invoke();}} System.Action h=()=>{}; var b=new Btn(); b.Click+=h; b.Click-=h; __P((b.Adds).ToString()); __P((b.Removes).ToString());
__Check("1\n1");
