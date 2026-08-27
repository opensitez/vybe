// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_3

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

IService_3 s = new DerivedService_3();
__P(s.GetName());
__Check("Service_3");

interface IService_3 {
    string GetName();
}
abstract class BaseService_3 : IService_3 {
    public abstract string GetName();
}
class DerivedService_3 : BaseService_3 {
    public override string GetName() => "Service_3";
}
