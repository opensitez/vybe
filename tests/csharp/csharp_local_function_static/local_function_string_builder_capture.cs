// vybe-test: csharp/csharp_local_function_static/local_function_string_builder_capture
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Join(int a,int b){var sb=new System.Text.StringBuilder(); string Append(int x){sb.Append(x); return sb.ToString();} Append(a); return Append(b);} __Check((Join(1,2)).ToString(), "12");
