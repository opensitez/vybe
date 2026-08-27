// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_10

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

IService_10 s = new DerivedService_10();
__P(s.GetName());
__Check("Service_10");

interface IService_10 {
    string GetName();
}
abstract class BaseService_10 : IService_10 {
    public abstract string GetName();
}
class DerivedService_10 : BaseService_10 {
    public override string GetName() => "Service_10";
}
