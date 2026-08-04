// vybe-test: csharp/csharp_array_advanced/array_convert_all_transforms_each_element
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

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

int[] src={1,2,3};
string[] dst=System.Array.ConvertAll(src,n=>n.ToString()+"x");
__P((dst[1]).ToString());
__Check("2x");
