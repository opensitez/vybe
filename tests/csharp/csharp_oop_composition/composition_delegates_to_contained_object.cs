// vybe-test: csharp/csharp_oop_composition/composition_delegates_to_contained_object
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

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

class Logger{public void Log(string m)=>__P(("[LOG]"+m).ToString());}
class Service{
    readonly Logger _log=new Logger();
    public void Do(string m){_log.Log(m);}
}
new Service().Do("hello");
__Check("[LOG]hello");
