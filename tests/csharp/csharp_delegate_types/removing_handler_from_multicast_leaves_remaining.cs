// vybe-test: csharp/csharp_delegate_types/removing_handler_from_multicast_leaves_remaining
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

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

int count = 0;
System.Action a = () => count++;
System.Action b = () => count++;
System.Action multi = a;
multi += b;
multi -= a;
multi();
__P((count).ToString());
__Check("1");
