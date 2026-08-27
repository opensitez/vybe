// vybe-test: csharp/csharp_diagnostics_stack_trace_frames/stack_trace_inspection_case_19

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

var trace = new System.Diagnostics.StackTrace();
__P((trace.FrameCount > 0).ToString());
__Check("True");
