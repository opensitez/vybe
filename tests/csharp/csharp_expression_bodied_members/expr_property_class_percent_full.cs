// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_percent_full
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

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

class Tank { public int level = 75; public int capacity = 100; public int Percent => level * 100 / capacity; }
__P((new Tank().Percent).ToString());
__Check("75");
