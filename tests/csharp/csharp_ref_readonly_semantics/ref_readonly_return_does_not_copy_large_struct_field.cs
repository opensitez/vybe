// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_return_does_not_copy_large_struct_field
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

struct Big{public int A; public int B; public int C;} Big item=new Big(); item.B=77; ref readonly int Read(ref Big target)=>ref target.B; __P((Read(ref item)).ToString());
__Check("77");
