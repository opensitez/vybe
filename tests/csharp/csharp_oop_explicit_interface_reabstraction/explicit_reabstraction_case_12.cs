// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_12

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

IService_12 s = new DerivedService_12();
__P(s.GetName());
__Check("Service_12");

interface IService_12 {
    string GetName();
}
abstract class BaseService_12 : IService_12 {
    public abstract string GetName();
}
class DerivedService_12 : BaseService_12 {
    public override string GetName() => "Service_12";
}
