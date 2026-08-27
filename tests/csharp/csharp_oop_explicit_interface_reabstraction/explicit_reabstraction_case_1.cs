// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_1

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

IService_1 s = new DerivedService_1();
__P(s.GetName());
__Check("Service_1");

interface IService_1 {
    string GetName();
}
abstract class BaseService_1 : IService_1 {
    public abstract string GetName();
}
class DerivedService_1 : BaseService_1 {
    public override string GetName() => "Service_1";
}
