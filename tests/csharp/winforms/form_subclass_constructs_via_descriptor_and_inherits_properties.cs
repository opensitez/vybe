// vybe-test: csharp/winforms/form_subclass_constructs_via_descriptor_and_inherits_properties
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MyForm : Form {
        }
        var f = new MyForm();
        f.Text = "hello";
        __Check((f.Text).ToString(), "hello");
        __Check((f.__control_type).ToString(), "Form");
