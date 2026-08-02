// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_void_initialize_method
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IInit<T> where T:IInit<T>{static abstract void Configure(T target);}
struct Config:IInit<Config>{public int Ready; public static void Configure(Config target){target.Ready=1;}}
var c=new Config(); Config.Configure(c); __Check((c.Ready).ToString(), "1");
