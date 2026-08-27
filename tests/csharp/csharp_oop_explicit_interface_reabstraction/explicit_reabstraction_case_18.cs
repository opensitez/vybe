// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_18

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

IService_18 s = new DerivedService_18();
__P(s.GetName());
__Check("Service_18");

interface IService_18 {
    string GetName();
}
abstract class BaseService_18 : IService_18 {
    public abstract string GetName();
}
class DerivedService_18 : BaseService_18 {
    public override string GetName() => "Service_18";
}
