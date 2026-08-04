// vybe-test: csharp/csharp_static_constructor_once/static_constructor_increments_counter_only_once_across_two_instance_allocations
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_once.rs

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

class Tracker {
    public static int Instances;
    static Tracker() { Instances++; }
}
_ = new Tracker();
_ = new Tracker();
__P((Tracker.Instances).ToString());
__Check("1");
