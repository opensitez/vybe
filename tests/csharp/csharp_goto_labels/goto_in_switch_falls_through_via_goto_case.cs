// vybe-test: csharp/csharp_goto_labels/goto_in_switch_falls_through_via_goto_case
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=1;
string r="";
switch(n){
    case 1: r+="one"; goto case 2;
    case 2: r+="two"; break;
    case 3: r+="three"; break;
}
__Check((r).ToString(), "onetwo");
