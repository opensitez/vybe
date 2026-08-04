// vybe-test: csharp/more_classes/object_initializer
// origin: languages/csharp/tests/csharp/test_more_classes.rs

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

class Config {
            public string host;
            public int port;
            public Config() {}
        }
        var c = new Config();
        c.host = "localhost";
        c.port = 8080;
        __P((c.host).ToString());
        __P((c.port).ToString());
__Check("localhost\n8080");
