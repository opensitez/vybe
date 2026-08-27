// vybe-test: csharp/csharp_diagnostics_metrics_meter_counters/meter_counter_case_17

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

var meter = new System.Diagnostics.Metrics.Meter("TestMeter_17", "1.0.0");
var counter = meter.CreateCounter<int>("requests_17");
counter.Add(10);
__P(meter.Name);
__P(counter.Name);
__Check("TestMeter_17\nrequests_17");
