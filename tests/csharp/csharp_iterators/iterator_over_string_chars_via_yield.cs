// vybe-test: csharp/csharp_iterators/iterator_over_string_chars_via_yield
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

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

System.Collections.Generic.IEnumerable<char> Vowels(string s) {
    foreach(char c in s) if("aeiou".Contains(c)) yield return c;
}
int count=0;
foreach(var _ in Vowels("hello world")) count++;
__P((count).ToString());
__Check("3");
