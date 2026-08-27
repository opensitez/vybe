// vybe-test: csharp/csharp_datetime_stopwatch_high_resolution/stopwatch_timestamp_case_10

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

long t1 = System.Diagnostics.Stopwatch.GetTimestamp();
long t2 = System.Diagnostics.Stopwatch.GetTimestamp();
__P((t2 >= t1).ToString());
__P((System.Diagnostics.Stopwatch.Frequency > 0).ToString());
__Check("True\nTrue");
