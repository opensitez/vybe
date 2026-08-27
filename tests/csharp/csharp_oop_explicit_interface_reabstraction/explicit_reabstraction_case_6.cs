// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_6

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

IService_6 s = new DerivedService_6();
__P(s.GetName());
__Check("Service_6");

interface IService_6 {
    string GetName();
}
abstract class BaseService_6 : IService_6 {
    public abstract string GetName();
}
class DerivedService_6 : BaseService_6 {
    public override string GetName() => "Service_6";
}
