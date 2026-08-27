// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_11

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

IService_11 s = new DerivedService_11();
__P(s.GetName());
__Check("Service_11");

interface IService_11 {
    string GetName();
}
abstract class BaseService_11 : IService_11 {
    public abstract string GetName();
}
class DerivedService_11 : BaseService_11 {
    public override string GetName() => "Service_11";
}
