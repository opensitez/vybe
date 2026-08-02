// vybe-test: csharp/basics/hello_world
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("Hello, World!").ToString(), "Hello, World!");
