#[test]
fn test_binding_source_datasource_roundtrip() {
    let code = r#"Partial Class Form1
    Inherits System.Windows.Forms.Form
    Friend WithEvents bs1 As System.Windows.Forms.BindingSource
    Friend WithEvents da1 As System.Windows.Forms.DataAdapter
    Friend WithEvents txtName As System.Windows.Forms.TextBox

    Private Sub InitializeComponent()
        Me.bs1 = New System.Windows.Forms.BindingSource()
        Me.da1 = New System.Windows.Forms.DataAdapter()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        Me.txtName.Location = New System.Drawing.Point(10, 10)
        Me.txtName.Size = New System.Drawing.Size(100, 25)
        Me.txtName.Text = ""
        Me.txtName.Name = "txtName"
        Me.txtName.DataBindings.Add("Text", Me.bs1, "Name")
        Me.txtName.TabIndex = 0
        Me.Controls.Add(Me.txtName)
        Me.bs1.DataSource = Me.da1
        Me.bs1.Name = "bs1"
        Me.da1.ConnectionString = "Data Source=test.db"
        Me.da1.SelectCommand = "SELECT * FROM users"
        Me.da1.Name = "da1"
        Me.ClientSize = New System.Drawing.Size(400, 300)
        Me.Text = "Test"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
    End Sub
End Class"#;

    let program = vybe_parser::parse_program(code).expect("Failed to parse");
    let cls = program.declarations.into_iter().find_map(|d| {
        if let vybe_parser::Declaration::Class(c) = d { Some(c) } else { None }
    }).expect("No class found");
    
    let form = vybe_forms::serialization::designer_parser::extract_form_from_designer(&cls)
        .expect("Failed to extract form");
    
    println!("Form: {} ({} controls)", form.name, form.controls.len());
    for ctrl in &form.controls {
        println!("  {} ({:?})", ctrl.name, ctrl.control_type);
        for (key, val) in ctrl.properties.iter() {
            println!("    {} = {:?}", key, val);
        }
    }
    
    // Check BindingSource has DataSource = da1
    let bs1 = form.controls.iter().find(|c| c.name == "bs1").expect("bs1 not found");
    let ds = bs1.properties.get_string("DataSource").expect("bs1 DataSource not found");
    assert_eq!(ds, "da1", "BindingSource DataSource should be da1");
    
    // Check DataAdapter has ConnectionString and SelectCommand
    let da1 = form.controls.iter().find(|c| c.name == "da1").expect("da1 not found");
    let cs = da1.properties.get_string("ConnectionString").expect("da1 ConnectionString not found");
    assert_eq!(cs, "Data Source=test.db");
    let sc = da1.properties.get_string("SelectCommand").expect("da1 SelectCommand not found");
    assert_eq!(sc, "SELECT * FROM users");
    
    // Check TextBox has DataBindings
    let txt = form.controls.iter().find(|c| c.name == "txtName").expect("txtName not found");
    let dbs = txt.properties.get_string("DataBindings.Source").expect("txtName DataBindings.Source not found");
    assert_eq!(dbs, "bs1");
    let dbt = txt.properties.get_string("DataBindings.Text").expect("txtName DataBindings.Text not found");
    assert_eq!(dbt, "Name");
    
    // Now round-trip: generate and re-parse
    let generated = vybe_forms::serialization::designer_codegen::generate_designer_code(&form);
    println!("\n=== Generated ===\n{}", generated);
    
    let program2 = vybe_parser::parse_program(&generated).expect("Failed to parse generated code");
    let cls2 = program2.declarations.into_iter().find_map(|d| {
        if let vybe_parser::Declaration::Class(c) = d { Some(c) } else { None }
    }).expect("No class in generated");
    let form2 = vybe_forms::serialization::designer_parser::extract_form_from_designer(&cls2)
        .expect("Failed to extract form from generated code");
    
    // Verify round-trip preserves BindingSource DataSource
    let bs1_2 = form2.controls.iter().find(|c| c.name == "bs1").expect("bs1 not found in roundtrip");
    let ds2 = bs1_2.properties.get_string("DataSource").expect("bs1 DataSource lost in roundtrip!");
    assert_eq!(ds2, "da1", "BindingSource DataSource should survive round-trip");
    
    // Verify round-trip preserves DataBindings
    let txt2 = form2.controls.iter().find(|c| c.name == "txtName").expect("txtName not found in roundtrip");
    let dbs2 = txt2.properties.get_string("DataBindings.Source").expect("DataBindings.Source lost in roundtrip!");
    assert_eq!(dbs2, "bs1");
    let dbt2 = txt2.properties.get_string("DataBindings.Text").expect("DataBindings.Text lost in roundtrip!");
    assert_eq!(dbt2, "Name");
}
