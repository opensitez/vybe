use super::helpers::run_csharp;

#[test]
fn new_button_has_properties() {
    let out = run_csharp(
        r#"
        var btn = new Button();
        btn.Name = "testBtn";
        btn.Text = "Click Me";
        Console.WriteLine(btn.Name);
        Console.WriteLine(btn.Text);
    "#,
    );
    assert_eq!(out, vec!["testBtn", "Click Me"]);
}

#[test]
fn fully_qualified_button_uses_shared_dotnet_surface() {
    let out = run_csharp(
        r#"
        var btn = new System.Windows.Forms.Button();
        btn.Text = "Shared";
        Console.WriteLine(btn.Text);
        Console.WriteLine(btn.__control_type);
    "#,
    );
    assert_eq!(out, vec!["Shared", "Button"]);
}

#[test]
fn new_point_properties() {
    let out = run_csharp(
        r#"
        var p = new Point(100, 200);
        Console.WriteLine(p.x);
        Console.WriteLine(p.y);
    "#,
    );
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn new_point_and_size() {
    let out = run_csharp(
        r#"
        var p = new Point(10, 20);
        var s = new Size(100, 50);
        Console.WriteLine(p.x + " " + p.y);
        Console.WriteLine(s.width + " " + s.height);
    "#,
    );
    assert_eq!(out, vec!["10 20", "100 50"]);
}

#[test]
fn winforms_button_with_event() {
    let out = run_csharp(
        r#"
        var btn = new Button();
        btn.Name = "btn1";
        btn.Text = "Click";
        Console.WriteLine(btn.Name);
        Console.WriteLine(btn.Text);
        Console.WriteLine(btn.__control_type);
    "#,
    );
    assert_eq!(out, vec!["btn1", "Click", "Button"]);
}

#[test]
fn winforms_application_bootstrap_methods_are_supported() {
    let out = run_csharp(
        r#"
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Console.WriteLine("boot-ok");
    "#,
    );
    assert_eq!(out, vec!["boot-ok"]);
}
