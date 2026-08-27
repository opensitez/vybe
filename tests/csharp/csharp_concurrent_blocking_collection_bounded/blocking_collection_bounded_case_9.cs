// vybe-test: csharp/csharp_concurrent_blocking_collection_bounded/blocking_collection_bounded_case_9

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

var bc = new System.Collections.Concurrent.BlockingCollection<int>(boundedCapacity: 2);
bc.Add(9);
bc.Add(10);
int t1 = bc.Take();
int t2 = bc.Take();
__P(t1.ToString());
__P(t2.ToString());
__Check("9\n10");
