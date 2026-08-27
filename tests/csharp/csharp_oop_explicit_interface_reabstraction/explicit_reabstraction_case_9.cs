// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_9

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

IService_9 s = new DerivedService_9();
__P(s.GetName());
__Check("Service_9");

interface IService_9 {
    string GetName();
}
abstract class BaseService_9 : IService_9 {
    public abstract string GetName();
}
class DerivedService_9 : BaseService_9 {
    public override string GetName() => "Service_9";
}
