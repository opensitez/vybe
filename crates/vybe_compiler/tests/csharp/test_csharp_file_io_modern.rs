use super::helpers::{compile_csharp_to_wasm, extract_imports, run_csharp};

fn temp_root() -> std::path::PathBuf {
    let mut dir = std::env::temp_dir();
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    dir.push(format!("vybe_csharp_file_io_{nanos}"));
    dir
}

fn cs_path(path: &std::path::Path) -> String {
    path.to_string_lossy().replace('\\', "\\\\")
}

#[test]
fn csharp_streamreader_and_streamwriter_use_real_surface_imports() {
    let root = temp_root();
    std::fs::create_dir_all(&root).unwrap();
    let file = root.join("stream.txt");

    let src = format!(
        r#"
using System.IO;

var writer = new StreamWriter("{file}");
writer.WriteLine("line-one");
writer.WriteLine("line-two");
writer.Close();

var reader = new StreamReader("{file}");
Console.WriteLine(reader.ReadLine());
Console.WriteLine(reader.ReadToEnd().Contains("line-two"));
reader.Close();
"#,
        file = cs_path(&file)
    );

    let output = run_csharp(&src);
    assert_eq!(output, vec!["line-one", "True"]);

    let wasm = compile_csharp_to_wasm(&src);
    let imports = extract_imports(&wasm);
    assert!(
        imports.iter().any(|(module, func)| module == "node:fs" && (func == "readFileSync" || func == "writeFileSync")),
        "expected StreamReader/StreamWriter emitter path to use node:fs imports, got {imports:?}"
    );
    assert!(
        imports.iter().all(|(module, _)| !module.starts_with("dotnet:")),
        "unexpected retired dotnet host import in {imports:?}"
    );

    let _ = std::fs::remove_dir_all(&root);
}