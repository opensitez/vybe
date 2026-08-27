// vybe-test: csharp/csharp_oop_unmanaged_callers_only_attribute/function_pointer_delegate_case_10

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

Func<int, int> f = x => x * 10;
__P(f(2).ToString());
__Check("20");
