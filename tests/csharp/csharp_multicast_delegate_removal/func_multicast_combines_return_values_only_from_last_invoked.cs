// vybe-test: csharp/csharp_multicast_delegate_removal/func_multicast_combines_return_values_only_from_last_invoked
// origin: languages/csharp/tests/csharp/test_csharp_multicast_delegate_removal.rs

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

using System;
Func<int> first = () => { __P(("1").ToString()); return 1; };
Func<int> second = () => { __P(("2").ToString()); return 2; };
Func<int> chain = first;
chain += second;
__P((chain()).ToString());
__Check("1\n2\n2");
