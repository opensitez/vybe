// vybe-test: csharp/csharp_oop_required_members_init/required_member_case_6

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

var p = new Person_6 { Name = "User_6" };
__P(p.Name);
__Check("User_6");

class Person_6 {
    public required string Name { get; init; }
}
