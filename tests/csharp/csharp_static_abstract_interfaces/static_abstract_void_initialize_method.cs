// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_void_initialize_method
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

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

interface IInit<T> where T:IInit<T>{static abstract void Configure(T target);}
struct Config:IInit<Config>{public int Ready; public static void Configure(Config target){target.Ready=1;}}
var c=new Config(); Config.Configure(c); __P((c.Ready).ToString());
__Check("1");
