// vybe-test: csharp/winforms/winforms_application_bootstrap_methods_are_supported
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        __Check(("boot-ok").ToString(), "boot-ok");
