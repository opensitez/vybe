// vybe-test: csharp/csharp_oop_composition/composition_delegates_to_contained_object
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Logger{public void Log(string m)=>__Check(("[LOG]"+m).ToString(), "[LOG]hello");}
class Service{
    readonly Logger _log=new Logger();
    public void Do(string m){_log.Log(m);}
}
new Service().Do("hello");
