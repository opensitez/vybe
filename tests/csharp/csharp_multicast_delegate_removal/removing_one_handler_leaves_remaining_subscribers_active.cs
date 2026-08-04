// vybe-test: csharp/csharp_multicast_delegate_removal/removing_one_handler_leaves_remaining_subscribers_active
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
void A() { __P(("A").ToString()); }
void B() { __P(("B").ToString()); }
Action chain = A;
chain += B;
chain -= A;
chain();
__Check("B");
