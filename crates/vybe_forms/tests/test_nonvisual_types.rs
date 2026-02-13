#[test]
fn test_nonvisual_types_roundtrip() {
    let code = r#"Partial Class TestForm
    Inherits System.Windows.Forms.Form
    Private Sub InitializeComponent()
        Me.bs1 = New System.Windows.Forms.BindingSource()
        Me.da1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.ds1 = New System.Data.DataSet()
        Me.dt1 = New System.Data.DataTable()
        Me.SuspendLayout()
        Me.bs1.Name = "bs1"
        Me.da1.Name = "da1"
        Me.ds1.Name = "ds1"
        Me.dt1.Name = "dt1"
        Me.ClientSize = New System.Drawing.Size(400, 300)
        Me.Text = "Test"
        Me.Name = "TestForm"
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
    assert_eq!(form.controls.len(), 4, "Should have 4 non-visual controls");
    
    for ctrl in &form.controls {
        println!("  {} -> {:?}", ctrl.name, ctrl.control_type);
    }

    // Check types
    let bs1 = form.controls.iter().find(|c| c.name == "bs1").expect("bs1 not found");
    assert!(matches!(bs1.control_type, vybe_forms::ControlType::BindingSourceComponent), "bs1 should be BindingSourceComponent, got {:?}", bs1.control_type);

    let da1 = form.controls.iter().find(|c| c.name == "da1").expect("da1 not found");
    assert!(matches!(da1.control_type, vybe_forms::ControlType::DataAdapterComponent), "da1 should be DataAdapterComponent, got {:?}", da1.control_type);

    let ds1 = form.controls.iter().find(|c| c.name == "ds1").expect("ds1 not found");
    assert!(matches!(ds1.control_type, vybe_forms::ControlType::DataSetComponent), "ds1 should be DataSetComponent, got {:?}", ds1.control_type);

    let dt1 = form.controls.iter().find(|c| c.name == "dt1").expect("dt1 not found");
    assert!(matches!(dt1.control_type, vybe_forms::ControlType::DataTableComponent), "dt1 should be DataTableComponent, got {:?}", dt1.control_type);

    // Now round-trip through codegen and re-parse
    let generated = vybe_forms::serialization::designer_codegen::generate_designer_code(&form);
    println!("\n=== Generated ===\n{}", generated);
    
    let program2 = vybe_parser::parse_program(&generated).expect("Failed to parse generated code");
    let cls2 = program2.declarations.into_iter().find_map(|d| {
        if let vybe_parser::Declaration::Class(c) = d { Some(c) } else { None }
    }).expect("No class in generated");
    let form2 = vybe_forms::serialization::designer_parser::extract_form_from_designer(&cls2)
        .expect("Failed to extract form from generated code");

    assert_eq!(form2.controls.len(), 4, "Round-trip should preserve 4 controls");
    
    for ctrl in &form2.controls {
        println!("  {} -> {:?} (round-trip)", ctrl.name, ctrl.control_type);
    }

    let bs1_2 = form2.controls.iter().find(|c| c.name == "bs1").expect("bs1 not found in roundtrip");
    assert!(matches!(bs1_2.control_type, vybe_forms::ControlType::BindingSourceComponent), "bs1 type lost in roundtrip: {:?}", bs1_2.control_type);
    
    let da1_2 = form2.controls.iter().find(|c| c.name == "da1").expect("da1 not found in roundtrip");
    assert!(matches!(da1_2.control_type, vybe_forms::ControlType::DataAdapterComponent), "da1 type lost in roundtrip: {:?}", da1_2.control_type);
}
