// vybe-test: csharp/csharp_ref_return_semantics/ref_return_from_local_function_updates_outer_variable
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

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

int total = 5;
ref int Bump() => ref total;
ref int view = ref Bump();
view += 2;
__P((total).ToString());
__Check("7");
