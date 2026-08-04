// vybe-test: csharp/csharp_object_initializers/with_expression_on_nominal_record_is_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

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

record Config{public int Port{get;init;}=80;}
var cfg=new Config() with{Port=443};
__P((cfg.Port).ToString());
__Check("443");
