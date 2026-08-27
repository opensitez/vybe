// vybe-test: csharp/more_classes/object_initializer
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

var c = new Config();
c.host = "localhost";
c.port = 8080;
__P((c.host).ToString());
__P((c.port).ToString());
__Check("localhost\n8080");

class Config {
            public string host;
            public int port;
            public Config() {}
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
