// vybe-test: csharp/csharp_oop_init_only_property_immutability/init_only_prop_case_10

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var cfg = new Config_10 { Timeout = 100 };
__P(cfg.Timeout.ToString());
__Check("100");

class Config_10 {
    public int Timeout { get; init; }
}
