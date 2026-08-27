// vybe-test: csharp/csharp_io_pipelines_pipe_reader_writer/pipelines_pipe_case_11

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

var pipe = new System.IO.Pipelines.Pipe();
__P((pipe.Reader != null).ToString());
__P((pipe.Writer != null).ToString());
__Check("True\nTrue");
