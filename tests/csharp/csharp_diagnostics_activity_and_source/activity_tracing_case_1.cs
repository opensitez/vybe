// vybe-test: csharp/csharp_diagnostics_activity_and_source/activity_tracing_case_1

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

var act = new System.Diagnostics.Activity("TestActivity_1");
act.SetTag("item.id", "1");
act.Start();
__P(act.OperationName);
__P(act.GetTagItem("item.id")?.ToString());
act.Stop();
__Check("TestActivity_1\n1");
