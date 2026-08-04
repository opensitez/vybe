// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_and_band_c
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

int n=55; __P((n switch{>=90=>"A",>=70 and <90=>"B",>=50 and <70=>"C",_=>"F"}).ToString());
__Check("C");
