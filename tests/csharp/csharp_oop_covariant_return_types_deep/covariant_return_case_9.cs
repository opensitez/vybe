// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_9

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

var repo = new DerivedRepo_9();
__P((repo.Get() is DerivedEntity_9).ToString());
__Check("True");

class BaseEntity_9 { }
class DerivedEntity_9 : BaseEntity_9 { }
abstract class BaseRepo_9 {
    public abstract BaseEntity_9 Get();
}
class DerivedRepo_9 : BaseRepo_9 {
    public override DerivedEntity_9 Get() => new DerivedEntity_9();
}
