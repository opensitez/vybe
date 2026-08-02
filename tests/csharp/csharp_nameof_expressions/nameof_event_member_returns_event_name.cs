// vybe-test: csharp/csharp_nameof_expressions/nameof_event_member_returns_event_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Publisher{public event System.Action Raised;} __Check((nameof(Publisher.Raised)).ToString(), "Raised");
