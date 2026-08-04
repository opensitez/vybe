// vybe-test: csharp/csharp_datetime_timespan/datetime_ticks_counts_interval_since_epoch_origin
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

var epoch = new System.DateTime(1970, 1, 1, 0, 0, 0, System.DateTimeKind.Utc); __P((epoch.Ticks > 0).ToString());
__Check("True");
