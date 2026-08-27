// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_14

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

IService_14 s = new DerivedService_14();
__P(s.GetName());
__Check("Service_14");

interface IService_14 {
    string GetName();
}
abstract class BaseService_14 : IService_14 {
    public abstract string GetName();
}
class DerivedService_14 : BaseService_14 {
    public override string GetName() => "Service_14";
}
