// vybe-test: csharp/csharp_oop_composition/composition_delegates_to_contained_object
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

using static __Harness;

new Service().Do("hello");
__Check("[LOG]hello");

class Logger{public void Log(string m)=>__P(("[LOG]"+m).ToString());}

class Service{
    readonly Logger _log=new Logger();
    public void Do(string m){_log.Log(m);}
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
