// vybe-test: csharp/exceptions_advanced/using_disposes_on_exception
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Conn : IDisposable {
    public void Dispose() { __Check(("conn closed").ToString(), "conn closed"); }
}
try {
    using (var c = new Conn()) {
        throw new Exception("fail");
    }
} catch (Exception e) {
    __Check(("caught: " + e.Message).ToString(), "caught: fail");
}
