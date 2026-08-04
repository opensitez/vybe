// vybe-test: csharp/csharp_datetime_timespan/datetime_kind_utc_survives_property_read
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

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

var instant = new System.DateTime(2024, 6, 1, 0, 0, 0, System.DateTimeKind.Utc); __P((instant.Kind).ToString());
__Check("Utc");
