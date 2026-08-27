// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_15

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

string res = FactoryCaller_15.Call();
__P(res);
__Check("Instance_15");

interface IFactory_15<TSelf> where TSelf : IFactory_15<TSelf> {
    static abstract string Create();
}
class FactoryImpl_15 : IFactory_15<FactoryImpl_15> {
    public static string Create() => "Instance_15";
}
class FactoryCaller_15 {
    public static string Call() => FactoryImpl_15.Create();
}
