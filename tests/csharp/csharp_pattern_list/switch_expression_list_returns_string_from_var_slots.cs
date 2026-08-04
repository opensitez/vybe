// vybe-test: csharp/csharp_pattern_list/switch_expression_list_returns_string_from_var_slots
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

string PairLabel(int[] a)=>a switch{[var x,var y]=>$"{x}-{y}",_=>"?"}; __P((PairLabel(new[]{7,8})).ToString());
__Check("7-8");
