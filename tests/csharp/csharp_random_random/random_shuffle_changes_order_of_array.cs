// vybe-test: csharp/csharp_random_random/random_shuffle_changes_order_of_array
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

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

int[] arr={1,2,3,4,5,6,7,8,9,10};
var rng=new System.Random(42);
rng.Shuffle(arr);
bool changed=false;
for(int i=0;i<arr.Length;i++) if(arr[i]!=i+1){changed=true;break;}
__P((changed).ToString());
__Check("True");
