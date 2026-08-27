// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_19

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

IService_19 s = new DerivedService_19();
__P(s.GetName());
__Check("Service_19");

interface IService_19 {
    string GetName();
}
abstract class BaseService_19 : IService_19 {
    public abstract string GetName();
}
class DerivedService_19 : BaseService_19 {
    public override string GetName() => "Service_19";
}
