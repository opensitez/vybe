// vybe-test: csharp/csharp_threading_cancellation_token_callbacks/cancellation_token_case_3

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

using var cts = new System.Threading.CancellationTokenSource();
bool canceled = false;
cts.Token.Register(() => canceled = true);
cts.Cancel();
__P(canceled.ToString());
__Check("True");
