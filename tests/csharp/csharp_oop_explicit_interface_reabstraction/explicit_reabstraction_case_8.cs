// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_8

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

IService_8 s = new DerivedService_8();
__P(s.GetName());
__Check("Service_8");

interface IService_8 {
    string GetName();
}
abstract class BaseService_8 : IService_8 {
    public abstract string GetName();
}
class DerivedService_8 : BaseService_8 {
    public override string GetName() => "Service_8";
}
