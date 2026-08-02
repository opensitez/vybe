// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_async_style_method_name
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Loader {
    public System.Threading.Tasks.Task<int> LoadAsync() {
        Trace.Show();
        return System.Threading.Tasks.Task.FromResult(1);
    }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "LoadAsync");
}
var t = new Loader().LoadAsync();
__Check((t.Result).ToString(), "1");
