// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_13

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

IService_13 s = new DerivedService_13();
__P(s.GetName());
__Check("Service_13");

interface IService_13 {
    string GetName();
}
abstract class BaseService_13 : IService_13 {
    public abstract string GetName();
}
class DerivedService_13 : BaseService_13 {
    public override string GetName() => "Service_13";
}
