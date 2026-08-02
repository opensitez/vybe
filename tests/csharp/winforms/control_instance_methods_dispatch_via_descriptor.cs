// vybe-test: csharp/winforms/control_instance_methods_dispatch_via_descriptor
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var b = new Button();
        b.Show();
        b.Hide();
        __Check(("methods-ok").ToString(), "methods-ok");
