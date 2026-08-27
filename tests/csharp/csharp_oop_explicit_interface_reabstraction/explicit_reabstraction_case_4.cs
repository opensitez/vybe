// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_4

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

IService_4 s = new DerivedService_4();
__P(s.GetName());
__Check("Service_4");

interface IService_4 {
    string GetName();
}
abstract class BaseService_4 : IService_4 {
    public abstract string GetName();
}
class DerivedService_4 : BaseService_4 {
    public override string GetName() => "Service_4";
}
