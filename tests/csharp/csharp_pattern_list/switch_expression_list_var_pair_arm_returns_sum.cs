// vybe-test: csharp/csharp_pattern_list/switch_expression_list_var_pair_arm_returns_sum
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

int SumPair(int[] a)=>a switch{[var x,var y]=>x+y,_=>0}; __P((SumPair(new[]{10,20})).ToString());
__Check("30");
