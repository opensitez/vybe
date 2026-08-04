// vybe-test: csharp/csharp_pattern_list/is_list_double_triple_sum_via_vars
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

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

double[] vals=new[]{1.5,2.0,2.5}; if(vals is [var a,var b,var c]) __P((a+b+c).ToString());
__Check("6");
