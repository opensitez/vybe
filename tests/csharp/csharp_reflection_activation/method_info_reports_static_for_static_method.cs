// vybe-test: csharp/csharp_reflection_activation/method_info_reports_static_for_static_method
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public static string Read() { return "ok"; } } __Check((typeof(Box).GetMethod("Read").IsStatic).ToString(), "True");
