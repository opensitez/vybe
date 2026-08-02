// vybe-test: csharp/csharp_exceptions_flow/finally_always_runs_even_after_return
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool ran=false;
int Compute(){
    try{return 42;}
    finally{ran=true;}
}
int v=Compute();
__Check((v).ToString(), "42"); __Check((ran).ToString(), "True");
