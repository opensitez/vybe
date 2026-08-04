// vybe-test: csharp/csharp_pattern_list/switch_statement_list_pattern_case_with_var_capture
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

int[] data=new[]{3,9}; string tag=""; switch(data){case[var a,var b]:tag=(a+b).ToString();break;default:tag="0";break;} __P((tag).ToString());
__Check("12");
