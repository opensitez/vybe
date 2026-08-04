// vybe-test: csharp/csharp_linq_advanced/append_adds_element_to_end_of_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

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

var result=new[]{1,2,3}.Append(4);
__P((result.Last()).ToString());
__Check("4");
