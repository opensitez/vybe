// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

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

// record_shape_basics
int seed = 39; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __P((numbers.Length == 3).ToString());
__Check("True");
