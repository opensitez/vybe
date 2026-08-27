// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_16

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

string res = FactoryCaller_16.Call();
__P(res);
__Check("Instance_16");

interface IFactory_16<TSelf> where TSelf : IFactory_16<TSelf> {
    static abstract string Create();
}
class FactoryImpl_16 : IFactory_16<FactoryImpl_16> {
    public static string Create() => "Instance_16";
}
class FactoryCaller_16 {
    public static string Call() => FactoryImpl_16.Create();
}
