// vybe-test: csharp/winforms/new_button_has_properties
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var btn = new Button();
        btn.Name = "testBtn";
        btn.Text = "Click Me";
        __Check((btn.Name).ToString(), "testBtn");
        __Check((btn.Text).ToString(), "Click Me");
