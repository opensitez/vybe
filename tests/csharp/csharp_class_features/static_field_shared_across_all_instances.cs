// vybe-test: csharp/csharp_class_features/static_field_shared_across_all_instances
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

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

class Ctr{public static int Count=0; public Ctr(){Count++;}}
new Ctr(); new Ctr(); new Ctr();
__P((Ctr.Count).ToString());
__Check("3");
