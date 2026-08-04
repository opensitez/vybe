// vybe-test: csharp/csharp_delegate_variance/func_covariant_two_independent_assignments
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

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

System.Func<string> a=()=>"one"; System.Func<string> b=()=>"two"; System.Func<object> ga=a; System.Func<object> gb=b; __P((ga()).ToString()); __P((gb()).ToString());
__Check("one\ntwo");
