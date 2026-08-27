// vybe-test: csharp/csharp_io_binary_reader_writer_primitives/binary_reader_writer_case_2

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

var ms = new System.IO.MemoryStream();
using (var bw = new System.IO.BinaryWriter(ms, System.Text.Encoding.UTF8, true)) {
    bw.Write(2);
    bw.Write("Text_2");
}
ms.Position = 0;
using (var br = new System.IO.BinaryReader(ms)) {
    int num = br.ReadInt32();
    string str = br.ReadString();
    __P(num.ToString());
    __P(str);
}
__Check("2\nText_2");
