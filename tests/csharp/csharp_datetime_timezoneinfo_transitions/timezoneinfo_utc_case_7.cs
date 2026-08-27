// vybe-test: csharp/csharp_datetime_timezoneinfo_transitions/timezoneinfo_utc_case_7

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

var utc = TimeZoneInfo.Utc;
__P(utc.Id);
__P(utc.BaseUtcOffset.TotalHours.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("UTC\n0");
