// vybe-test: csharp/csharp_nullable_ref_types/nullable_reference_type_array_element_can_be_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

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

string?[] arr=new string?[3];
arr[0]="a"; arr[1]=null; arr[2]="c";
int nonNull=0;
foreach(var s in arr) if(s!=null) nonNull++;
__P((nonNull).ToString());
__Check("2");
