// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_7

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

IService_7 s = new DerivedService_7();
__P(s.GetName());
__Check("Service_7");

interface IService_7 {
    string GetName();
}
abstract class BaseService_7 : IService_7 {
    public abstract string GetName();
}
class DerivedService_7 : BaseService_7 {
    public override string GetName() => "Service_7";
}
