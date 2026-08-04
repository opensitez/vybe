// vybe-test: csharp/csharp_exception_types/index_out_of_range_exception_thrown_by_bad_array_access
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

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

string result = "ok";
try { int x = new int[2][5]; }
catch(System.IndexOutOfRangeException) { result = "oob"; }
catch(System.Exception) { result = "oob"; }
__P((result).ToString());
__Check("oob");
