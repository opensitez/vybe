// vybe-test: csharp/winforms/winforms_button_with_event
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var btn = new Button();
        btn.Name = "btn1";
        btn.Text = "Click";
        __Check((btn.Name).ToString(), "btn1");
        __Check((btn.Text).ToString(), "Click");
        __Check((btn.__control_type).ToString(), "Button");
