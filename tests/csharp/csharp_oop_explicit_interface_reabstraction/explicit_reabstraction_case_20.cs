// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_20

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

IService_20 s = new DerivedService_20();
__P(s.GetName());
__Check("Service_20");

interface IService_20 {
    string GetName();
}
abstract class BaseService_20 : IService_20 {
    public abstract string GetName();
}
class DerivedService_20 : BaseService_20 {
    public override string GetName() => "Service_20";
}
