// vybe-test: csharp/csharp_disposable_pattern/using_declaration_disposes_at_end_of_block
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{public bool Gone;public void Dispose(){Gone=true;}}
R r;
{using var x=new R(); r=x;}
__Check((r.Gone).ToString(), "True");
