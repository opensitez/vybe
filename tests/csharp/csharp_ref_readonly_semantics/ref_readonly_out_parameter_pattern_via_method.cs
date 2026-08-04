// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_out_parameter_pattern_via_method
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

bool TryGet(ref readonly int[] src,int i,out int value){value=src[i]; return true;} int[] arr={12}; TryGet(ref arr,0,out int v); __P((v).ToString());
__Check("12");
