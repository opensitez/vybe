// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_5

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

IService_5 s = new DerivedService_5();
__P(s.GetName());
__Check("Service_5");

interface IService_5 {
    string GetName();
}
abstract class BaseService_5 : IService_5 {
    public abstract string GetName();
}
class DerivedService_5 : BaseService_5 {
    public override string GetName() => "Service_5";
}
