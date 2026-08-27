// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_13

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

var repo = new DerivedRepo_13();
__P((repo.Get() is DerivedEntity_13).ToString());
__Check("True");

class BaseEntity_13 { }
class DerivedEntity_13 : BaseEntity_13 { }
abstract class BaseRepo_13 {
    public abstract BaseEntity_13 Get();
}
class DerivedRepo_13 : BaseRepo_13 {
    public override DerivedEntity_13 Get() => new DerivedEntity_13();
}
