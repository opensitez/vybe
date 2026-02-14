use vybe_forms::serialization::designer_parser;
use vybe_forms::serialization::designer_codegen;
use vybe_parser::parse_program;

#[test]
fn test_designer_arbitrary_properties() {
    let vb_code = r#"
Partial Class TestForm
    Inherits System.Windows.Forms.Form
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        
        ' Button1 Standard Props
        Me.Button1.Name = "Button1"
        Me.Button1.Text = "Click Me"
        
        ' Button1 Arbitrary Props
        Me.Button1.Enabled = False
        Me.Button1.Visible = True
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Image = Global.My.Resources.Resources.MyImage
        
        ' Label1 Arbitrary Props
        Me.Label1.Name = "Label1"
        Me.Label1.Tag = "MyTag"
        Me.Label1.AutoSize = True
        Me.Label1.CustomProp = 123
        
        Me.ResumeLayout(False)
    End Sub
End Class
"#;

    let program = parse_program(vb_code).expect("Failed to parse VB code");
    let class_decl = program.declarations.iter().find_map(|d| match d {
        vybe_parser::Declaration::Class(c) => Some(c),
        _ => None,
    }).expect("No class found");
    
    // 1. Extract
    let form = designer_parser::extract_form_from_designer(class_decl)
        .expect("Failed to extract form");
        
    // 2. Verify Properties in Memory
    let btn = form.controls.iter().find(|c| c.name == "Button1").expect("Button1 missing");
    assert_eq!(btn.properties.get_bool("Enabled"), Some(false));
    assert_eq!(btn.properties.get_bool("Visible"), Some(true));
    // FlatStyle should be expression or something? "System.Windows.Forms.FlatStyle.Flat"
    // Let's check raw value type
    let flat_style = btn.properties.get("FlatStyle").expect("FlatStyle missing");
    if let vybe_forms::properties::PropertyValue::Expression(code) = flat_style {
        assert_eq!(code, "System.Windows.Forms.FlatStyle.Flat");
    } else {
        panic!("FlatStyle should be an Expression, got {:?}", flat_style);
    }
    
    let lbl = form.controls.iter().find(|c| c.name == "Label1").expect("Label1 missing");
    assert_eq!(lbl.properties.get_string("Tag"), Some("MyTag"));
    assert_eq!(lbl.properties.get_bool("AutoSize"), Some(true));
    assert_eq!(lbl.properties.get_int("CustomProp"), Some(123));
    
    // 3. Generate
    let generated_code = designer_codegen::generate_designer_code(&form);
    
    // 4. Verify Generated Code (Round Trip)
    println!("Generated Code:\n{}", generated_code);
    
    assert!(generated_code.contains("Me.Button1.Enabled = False"));
    assert!(generated_code.contains("Me.Button1.Visible = True"));
    assert!(generated_code.contains("Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat"));
    // Global.My... might be complex expression
    assert!(generated_code.contains("Me.Button1.Image = Global.My.Resources.Resources.MyImage"));
    
    assert!(generated_code.contains("Me.Label1.Tag = \"MyTag\"")); // Tag is special-cased in parse but not generic in codegen? No, it's blacklisted in codegen but not written?
    // Wait, Tag IS blacklisted in codegen. 
    // And codegen only writes Tag if array index.
    // So if I set arbitrary Tag, codegen might SKIP it if I blacklisted "Tag".
    // I need to check this behavior.
    
    assert!(generated_code.contains("Me.Label1.AutoSize = True"));
    assert!(generated_code.contains("Me.Label1.CustomProp = 123"));
}
