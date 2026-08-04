// vybe-test: csharp/csharp_ref_out_in/out_parameter_initialised_inside_callee
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

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

void Minmax(int[] a, out int min, out int max){
    min=a[0]; max=a[0];
    foreach(var v in a){if(v<min)min=v; if(v>max)max=v;}
}
Minmax(new[]{3,1,4,1,5,9}, out int lo, out int hi);
__P((lo).ToString()); __P((hi).ToString());
__Check("1\n9");
