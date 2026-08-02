// vybe-test: csharp/more_classes/object_initializer
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
        __Check((c.host).ToString(), "localhost");
        __Check((c.port).ToString(), "8080");
