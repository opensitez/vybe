// vybe-test: csharp/csharp_datetime_formatting/tostring_d_short_date_pattern_contains_year_digits
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

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

var d = new System.DateTime(2025,12,31);
__P((d.ToString("yyyy-MM-dd").StartsWith("2025")).ToString());
__Check("True");
