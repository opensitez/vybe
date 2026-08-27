// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_15

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

IService_15 s = new DerivedService_15();
__P(s.GetName());
__Check("Service_15");

interface IService_15 {
    string GetName();
}
abstract class BaseService_15 : IService_15 {
    public abstract string GetName();
}
class DerivedService_15 : BaseService_15 {
    public override string GetName() => "Service_15";
}
