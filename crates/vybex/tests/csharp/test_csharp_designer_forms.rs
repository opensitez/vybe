#[test]
fn csharp_designer_loads_controls_and_event_bindings() {
    let src = r#"
public partial class Form1 : System.Windows.Forms.Form
{
    private System.Windows.Forms.Button btn1;

    private void InitializeComponent()
    {
        this.btn1 = new System.Windows.Forms.Button();
        this.btn1.Location = new System.Drawing.Point(140, 140);
        this.btn1.Size = new System.Drawing.Size(120, 30);
        this.btn1.Text = "Hello";
        this.btn1.Name = "btn1";
        this.btn1.Click += this.btn1_Click;
        this.Controls.Add(this.btn1);
    }
}
"#;

    let mut gui = vybe_host::GuiState::new();
    vybex::languages::csharp::forms::load_designer(src, &mut gui)
        .expect("C# designer load should succeed");

    assert!(gui.control_names.iter().any(|n| n == "btn1"));
    assert_eq!(gui.get_property("btn1", "Text"), "Hello");

    let event_keys = gui.event_keys();
    assert!(
        event_keys.iter().any(|k| k == "btn1.click"),
        "expected btn1.click in event keys, got {:?}",
        event_keys
    );
}

#[test]
fn csharp_designer_file_lowers_click_subscriptions_to_addhandler() {
    use vybex::ast::{ClassMember, StmtKind};

    let src = include_str!("../../../../examples/csharp/cosmic_arcade_designer/MainForm.Designer.cs");
    let module = vybex::languages::csharp::parse(src).expect("C# parse failed");

    let mut init_stmt_kinds = Vec::new();
    let has_addhandler = module.body.iter().any(|stmt| {
        let StmtKind::ClassDecl { members, .. } = &stmt.kind else {
            return false;
        };

        members.iter().any(|member| {
            let ClassMember::Method(method) = member else {
                return false;
            };
            let StmtKind::FunctionDecl { name, body, .. } = &method.kind else {
                return false;
            };

            if name == "InitializeComponent" {
                init_stmt_kinds = body.iter().map(|stmt| format!("{:?}", stmt.kind)).collect();
                body.iter().any(|stmt| matches!(stmt.kind, StmtKind::AddHandler { .. }))
            } else {
                false
            }
        })
    });

    assert!(
        has_addhandler,
        "expected InitializeComponent in real designer file to contain AddHandler; got {:?}",
        init_stmt_kinds
    );
}

#[test]
fn csharp_event_unsubscription_lowers_to_removehandler() {
    use vybex::ast::{ClassMember, StmtKind};

    let src = r#"
public partial class Form1 : System.Windows.Forms.Form
{
    private System.Windows.Forms.Button btn1;

    private void InitializeComponent()
    {
        this.btn1.Click -= this.btn1_Click;
    }
}
"#;

    let module = vybex::languages::csharp::parse(src).expect("C# parse failed");
    let has_removehandler = module.body.iter().any(|stmt| {
        let StmtKind::ClassDecl { members, .. } = &stmt.kind else {
            return false;
        };

        members.iter().any(|member| {
            let ClassMember::Method(method) = member else {
                return false;
            };
            let StmtKind::FunctionDecl { name, body, .. } = &method.kind else {
                return false;
            };

            name == "InitializeComponent"
                && body.iter().any(|stmt| matches!(stmt.kind, StmtKind::RemoveHandler { .. }))
        })
    });

    assert!(
        has_removehandler,
        "expected event unsubscription to lower to RemoveHandler"
    );
}
