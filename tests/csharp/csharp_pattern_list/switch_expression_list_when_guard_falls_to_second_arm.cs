// vybe-test: csharp/csharp_pattern_list/switch_expression_list_when_guard_falls_to_second_arm
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

string Rank(int[] a)=>a switch{[var x,var y] when x>y=>"desc",[var x,var y]=>"asc",_=>"other"}; __P((Rank(new[]{2,5})).ToString());
__Check("asc");
