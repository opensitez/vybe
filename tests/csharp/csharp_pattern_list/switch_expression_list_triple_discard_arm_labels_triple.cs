// vybe-test: csharp/csharp_pattern_list/switch_expression_list_triple_discard_arm_labels_triple
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

string Label(int[] a)=>a switch{[_,_,_]=>"triple",_=>"other"}; __P((Label(new[]{1,2,3})).ToString());
__Check("triple");
