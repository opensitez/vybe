// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_17

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

IService_17 s = new DerivedService_17();
__P(s.GetName());
__Check("Service_17");

interface IService_17 {
    string GetName();
}
abstract class BaseService_17 : IService_17 {
    public abstract string GetName();
}
class DerivedService_17 : BaseService_17 {
    public override string GetName() => "Service_17";
}
