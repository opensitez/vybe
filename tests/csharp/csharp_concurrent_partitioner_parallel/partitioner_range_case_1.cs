// vybe-test: csharp/csharp_concurrent_partitioner_parallel/partitioner_range_case_1

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

var partitioner = System.Collections.Concurrent.Partitioner.Create(0, 10, 5);
var ranges = partitioner.GetOrderablePartitions(2);
__P((ranges != null).ToString());
__Check("True");
