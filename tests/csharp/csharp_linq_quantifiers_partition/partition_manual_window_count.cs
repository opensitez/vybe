// vybe-test: csharp/csharp_linq_quantifiers_partition/partition_manual_window_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

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

var src=new[]{10,20,30,40,50};
int size=2;
int windows=0;
for(int i=0;i+size<=src.Length;i+=size) windows++;
__P((windows).ToString());
__Check("2");
