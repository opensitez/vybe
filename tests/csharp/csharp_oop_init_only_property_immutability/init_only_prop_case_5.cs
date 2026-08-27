// vybe-test: csharp/csharp_oop_init_only_property_immutability/init_only_prop_case_5

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

var cfg = new Config_5 { Timeout = 50 };
__P(cfg.Timeout.ToString());
__Check("50");

class Config_5 {
    public int Timeout { get; init; }
}
