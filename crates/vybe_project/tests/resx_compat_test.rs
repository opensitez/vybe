use vybe_project::resources::{ResourceManager, ResourceItem, ResourceType};

#[test]
fn test_resx_round_trip_typed_resources() {
    let mut mgr = ResourceManager::new();
    mgr.resources.push(ResourceItem::new_string("AppTitle", "My Application"));
    mgr.resources.push(ResourceItem::new_file("logo", "Resources\\logo.png", ResourceType::Image));
    mgr.resources.push(ResourceItem::new_file("appIcon", "Resources\\app.ico", ResourceType::Icon));
    mgr.resources.push(ResourceItem::new_file("alert", "Resources\\alert.wav", ResourceType::Audio));
    mgr.resources.push(ResourceItem::new_file("readme", "Resources\\readme.txt", ResourceType::File));

    let resx = mgr.to_resx().unwrap();
    println!("Generated .resx:\n{}", resx);

    // Verify XML declaration
    assert!(resx.starts_with("<?xml version=\"1.0\" encoding=\"utf-8\"?>"));

    // Verify resheader elements
    assert!(resx.contains("text/microsoft-resx"));
    assert!(resx.contains("ResXResourceReader"));
    assert!(resx.contains("ResXResourceWriter"));

    // Verify file refs use semicolon format
    assert!(resx.contains("Resources\\logo.png;System.Drawing.Bitmap"));
    assert!(resx.contains("Resources\\app.ico;System.Drawing.Icon"));
    assert!(resx.contains("Resources\\alert.wav;System.IO.MemoryStream"));
    assert!(resx.contains("Resources\\readme.txt;System.Byte[]"));

    // String resource should NOT have ResXFileRef type
    assert!(resx.contains("<data name=\"AppTitle\""));

    // Round-trip parse
    let parsed = ResourceManager::parse_resx(&resx).unwrap();
    assert_eq!(parsed.resources.len(), 5);

    assert_eq!(parsed.resources[0].name, "AppTitle");
    assert_eq!(parsed.resources[0].resource_type, ResourceType::String);
    assert_eq!(parsed.resources[0].value, "My Application");

    assert_eq!(parsed.resources[1].name, "logo");
    assert_eq!(parsed.resources[1].resource_type, ResourceType::Image);
    assert_eq!(parsed.resources[1].file_name, Some("Resources\\logo.png".to_string()));

    assert_eq!(parsed.resources[2].name, "appIcon");
    assert_eq!(parsed.resources[2].resource_type, ResourceType::Icon);
    assert_eq!(parsed.resources[2].file_name, Some("Resources\\app.ico".to_string()));

    assert_eq!(parsed.resources[3].name, "alert");
    assert_eq!(parsed.resources[3].resource_type, ResourceType::Audio);

    assert_eq!(parsed.resources[4].name, "readme");
    assert_eq!(parsed.resources[4].resource_type, ResourceType::File);
}

#[test]
fn test_parse_real_vs_resx() {
    // A simplified but structurally accurate VS-generated .resx
    let vs_resx = r#"<?xml version="1.0" encoding="utf-8"?>
<root>
  <resheader name="resmimetype">
    <value>text/microsoft-resx</value>
  </resheader>
  <resheader name="version">
    <value>2.0</value>
  </resheader>
  <resheader name="reader">
    <value>System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>
  </resheader>
  <resheader name="writer">
    <value>System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>
  </resheader>
  <data name="WelcomeMessage" xml:space="preserve">
    <value>Welcome to the application!</value>
    <comment>Shown on startup</comment>
  </data>
  <data name="Logo" type="System.Resources.ResXFileRef, System.Windows.Forms">
    <value>..\Resources\logo.png;System.Drawing.Bitmap, System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a</value>
  </data>
  <data name="AppIcon" type="System.Resources.ResXFileRef, System.Windows.Forms">
    <value>..\Resources\app.ico;System.Drawing.Icon, System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a</value>
  </data>
  <data name="AlertSound" type="System.Resources.ResXFileRef, System.Windows.Forms">
    <value>..\Resources\alert.wav;System.IO.MemoryStream, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>
  </data>
</root>"#;

    let parsed = ResourceManager::parse_resx(vs_resx).unwrap();
    assert_eq!(parsed.resources.len(), 4);

    // String resource
    assert_eq!(parsed.resources[0].name, "WelcomeMessage");
    assert_eq!(parsed.resources[0].resource_type, ResourceType::String);
    assert_eq!(parsed.resources[0].value, "Welcome to the application!");
    assert_eq!(parsed.resources[0].comment, Some("Shown on startup".to_string()));

    // Image (ResXFileRef)
    assert_eq!(parsed.resources[1].name, "Logo");
    assert_eq!(parsed.resources[1].resource_type, ResourceType::Image);
    assert_eq!(parsed.resources[1].value, "..\\Resources\\logo.png");
    assert_eq!(parsed.resources[1].file_name, Some("..\\Resources\\logo.png".to_string()));

    // Icon (ResXFileRef)
    assert_eq!(parsed.resources[2].name, "AppIcon");
    assert_eq!(parsed.resources[2].resource_type, ResourceType::Icon);
    assert_eq!(parsed.resources[2].file_name, Some("..\\Resources\\app.ico".to_string()));

    // Audio (ResXFileRef → MemoryStream)
    assert_eq!(parsed.resources[3].name, "AlertSound");
    assert_eq!(parsed.resources[3].resource_type, ResourceType::Audio);
    assert_eq!(parsed.resources[3].file_name, Some("..\\Resources\\alert.wav".to_string()));
}
