// vybe-test: csharp/csharp_nameof_expressions/nameof_public_field_on_type_returns_field_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Account{public int Balance;} __Check((nameof(Account.Balance)).ToString(), "Balance");
