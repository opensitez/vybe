// vybe-test: csharp/winforms/fully_qualified_button_uses_shared_dotnet_surface
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var btn = new System.Windows.Forms.Button();
        btn.Text = "Shared";
        __Check((btn.Text).ToString(), "Shared");
        __Check((btn.__control_type).ToString(), "Button");
