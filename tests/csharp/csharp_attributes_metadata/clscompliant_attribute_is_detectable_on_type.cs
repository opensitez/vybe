// vybe-test: csharp/csharp_attributes_metadata/clscompliant_attribute_is_detectable_on_type
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using System; [CLSCompliant(true)] class PublicApi { } __P((Attribute.IsDefined(typeof(PublicApi), typeof(CLSCompliantAttribute))).ToString());
__Check("True");
