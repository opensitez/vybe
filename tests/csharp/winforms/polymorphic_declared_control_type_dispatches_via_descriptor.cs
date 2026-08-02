// vybe-test: csharp/winforms/polymorphic_declared_control_type_dispatches_via_descriptor
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Control c = new Button();
        c.Text = "poly";
        c.Show();
        __Check((c.Text).ToString(), "poly");
