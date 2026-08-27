// vybe-test: csharp/csharp_io_string_reader_writer_buffers/string_reader_writer_case_8

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

using var sw = new System.IO.StringWriter();
sw.WriteLine("Line1_8");
sw.WriteLine("Line2_8");
using var sr = new System.IO.StringReader(sw.ToString());
string l1 = sr.ReadLine();
string l2 = sr.ReadLine();
__P(l1);
__P(l2);
__Check("Line1_8\nLine2_8");
