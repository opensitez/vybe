// vybe-test: csharp/csharp_reflection_method_create_delegate/create_delegate_case_1

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

var method = typeof(Math).GetMethod("Abs", new Type[] { typeof(int) });
var func = (Func<int, int>)method.CreateDelegate(typeof(Func<int, int>));
int res = func(-1);
__P(res.ToString());
__Check("1");
