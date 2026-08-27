// vybe-test: csharp/csharp_oop_explicit_interface_reabstraction/explicit_reabstraction_case_2

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

IService_2 s = new DerivedService_2();
__P(s.GetName());
__Check("Service_2");

interface IService_2 {
    string GetName();
}
abstract class BaseService_2 : IService_2 {
    public abstract string GetName();
}
class DerivedService_2 : BaseService_2 {
    public override string GetName() => "Service_2";
}
