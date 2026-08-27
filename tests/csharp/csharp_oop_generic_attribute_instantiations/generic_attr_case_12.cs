// vybe-test: csharp/csharp_oop_generic_attribute_instantiations/generic_attr_case_12

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

var attr = (ValidatorAttribute_12<int>)typeof(TargetClass_12).GetCustomAttributes(typeof(ValidatorAttribute_12<int>), false)[0];
__P(attr.TargetType.Name);
__Check("Int32");

[System.AttributeUsage(System.AttributeTargets.Class)]
class ValidatorAttribute_12<T> : System.Attribute {
    public Type TargetType => typeof(T);
}
[ValidatorAttribute_12<int>]
class TargetClass_12 { }
