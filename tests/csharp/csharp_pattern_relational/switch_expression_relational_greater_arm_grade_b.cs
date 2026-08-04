// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_greater_arm_grade_b
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

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

int score=85; __P((score switch{>=90=>"A",>=80=>"B",_=>"C"}).ToString());
__Check("B");
