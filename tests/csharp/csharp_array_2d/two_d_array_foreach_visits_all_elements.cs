// vybe-test: csharp/csharp_array_2d/two_d_array_foreach_visits_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

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

int[,] m={{1,2},{3,4}};
int sum=0; foreach(int n in m) sum+=n;
__P((sum).ToString());
__Check("10");
