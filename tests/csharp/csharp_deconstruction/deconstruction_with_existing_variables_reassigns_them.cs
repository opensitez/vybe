// vybe-test: csharp/csharp_deconstruction/deconstruction_with_existing_variables_reassigns_them
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

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

int first = 0;
int second = 0;
(first, second) = (7, 9);
__P((first).ToString());
__P((second).ToString());
__Check("7\n9");
