// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_16

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

IService_16 s = new DerivedService_16();
__P(s.GetName());
__Check("Service_16");

interface IService_16 {
    string GetName();
}
abstract class BaseService_16 : IService_16 {
    public abstract string GetName();
}
class DerivedService_16 : BaseService_16 {
    public override string GetName() => "Service_16";
}
